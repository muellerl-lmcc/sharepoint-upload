import logging
import os
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import List

from sharepoint_sync.change_detector import ChangeDetector, FileStatus, compute_file_fingerprint
from sharepoint_sync.config import AppConfig
from sharepoint_sync.folder_walker import FolderWalker
from sharepoint_sync.sharepoint_client import SharePointClient
from sharepoint_sync.state_manager import FileRecord, StateManager

log = logging.getLogger(__name__)


@dataclass
class SyncReport:
    uploaded: int = 0
    skipped: int = 0
    deleted: int = 0
    errors: List[str] = field(default_factory=list)


class Uploader:
    def __init__(self, config: AppConfig, sp_client: SharePointClient, state: StateManager) -> None:
        self._config = config
        self._sp = sp_client
        self._state = state
        self._detector = ChangeDetector(state)
        self._walker = FolderWalker(config.source_folder)

    def run(self, dry_run: bool = False, force_full_sync: bool = False) -> SyncReport:
        report = SyncReport()
        seen_paths: set[str] = set()

        for abs_path, rel_path in self._walker.walk():
            seen_paths.add(rel_path)
            try:
                if force_full_sync:
                    status = FileStatus.NEW
                    fingerprint = None
                else:
                    status, fingerprint = self._detector.classify(abs_path, rel_path)

                if status == FileStatus.UNCHANGED:
                    report.skipped += 1
                    continue

                if dry_run:
                    log.info("[DRY-RUN] Would upload: %s (%s)", rel_path, status.name)
                    report.uploaded += 1
                    continue

                # Ensure parent folder exists
                parent = "/".join(rel_path.split("/")[:-1])
                if parent:
                    self._sp.ensure_folder(parent)

                item_id = self._sp.upload_file(abs_path, rel_path)

                mtime, size, md5 = fingerprint if fingerprint else compute_file_fingerprint(abs_path)
                record = FileRecord(
                    relative_path=rel_path,
                    mtime=mtime,
                    size=size,
                    md5=md5,
                    uploaded_at=datetime.now(timezone.utc).isoformat(),
                    sharepoint_item_id=item_id,
                )
                self._state.update(record)
                report.uploaded += 1

            except Exception as e:
                log.error("Error processing %s: %s", rel_path, e)
                report.errors.append(f"{rel_path}: {e}")

        if self._config.sync.delete_remote_if_local_deleted:
            for tracked_path in self._state.all_paths() - seen_paths:
                record = self._state.get(tracked_path)
                if record and record.sharepoint_item_id:
                    try:
                        if dry_run:
                            log.info("[DRY-RUN] Would delete remote: %s", tracked_path)
                        else:
                            self._sp.delete_item(record.sharepoint_item_id)
                            self._state.remove(tracked_path)
                        report.deleted += 1
                    except Exception as e:
                        log.error("Error deleting %s: %s", tracked_path, e)
                        report.errors.append(f"DELETE {tracked_path}: {e}")

        if not dry_run:
            self._state.save()

        log.info(
            "Sync complete — uploaded: %d, skipped: %d, deleted: %d, errors: %d",
            report.uploaded, report.skipped, report.deleted, len(report.errors),
        )
        return report
