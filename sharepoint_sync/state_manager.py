import json
import logging
import os
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from typing import Optional

log = logging.getLogger(__name__)

SCHEMA_VERSION = 1


@dataclass
class FileRecord:
    relative_path: str
    mtime: float
    size: int
    md5: str
    uploaded_at: str
    sharepoint_item_id: str


class StateManager:
    def __init__(self, state_file: str) -> None:
        self._path = state_file
        self._records: dict[str, FileRecord] = {}
        self._load()

    def _load(self) -> None:
        if not os.path.exists(self._path):
            return
        try:
            with open(self._path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for rel_path, rec in data.get("files", {}).items():
                self._records[rel_path] = FileRecord(**rec)
            log.debug("Loaded state: %d records", len(self._records))
        except Exception as e:
            log.warning("Could not load state file, starting fresh: %s", e)

    def get(self, relative_path: str) -> Optional[FileRecord]:
        return self._records.get(relative_path)

    def update(self, record: FileRecord) -> None:
        self._records[record.relative_path] = record

    def remove(self, relative_path: str) -> None:
        self._records.pop(relative_path, None)

    def all_paths(self) -> set[str]:
        return set(self._records.keys())

    def save(self) -> None:
        data = {
            "schema_version": SCHEMA_VERSION,
            "last_run": datetime.now(timezone.utc).isoformat(),
            "files": {k: asdict(v) for k, v in self._records.items()},
        }
        tmp = self._path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self._path)
        log.debug("State saved: %d records", len(self._records))
