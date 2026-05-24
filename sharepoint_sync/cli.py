import argparse
import logging
import sys

log = logging.getLogger(__name__)


def _run_status(config, source_override: str | None) -> None:
    """Local-only change detection — no Azure connection required."""
    from sharepoint_sync.change_detector import ChangeDetector, FileStatus
    from sharepoint_sync.folder_walker import FolderWalker
    from sharepoint_sync.state_manager import StateManager

    if source_override:
        config.source_folder = source_override

    state = StateManager(config.state.state_file)
    walker = FolderWalker(config.source_folder)
    detector = ChangeDetector(state)

    counts = {FileStatus.NEW: [], FileStatus.MODIFIED: [], FileStatus.UNCHANGED: []}

    for abs_path, rel_path in walker.walk():
        status, _ = detector.classify(abs_path, rel_path)
        counts[status].append(rel_path)

    # Files tracked in state but no longer on disk
    import os
    local_paths = {rel for _, rel in FolderWalker(config.source_folder).walk()}
    deleted = sorted(state.all_paths() - local_paths)

    print(f"\nStatus: {config.source_folder}")
    print(f"  Neu:        {len(counts[FileStatus.NEW])}")
    print(f"  Geändert:   {len(counts[FileStatus.MODIFIED])}")
    print(f"  Unverändert:{len(counts[FileStatus.UNCHANGED])}")
    print(f"  Lokal gelöscht: {len(deleted)}")

    if counts[FileStatus.NEW]:
        print("\nNeue Dateien:")
        for p in sorted(counts[FileStatus.NEW]):
            print(f"  + {p}")
    if counts[FileStatus.MODIFIED]:
        print("\nGeänderte Dateien:")
        for p in sorted(counts[FileStatus.MODIFIED]):
            print(f"  ~ {p}")
    if deleted:
        print("\nLokal gelöscht (noch in State):")
        for p in deleted:
            print(f"  - {p}")

    needs_sync = len(counts[FileStatus.NEW]) + len(counts[FileStatus.MODIFIED]) + len(deleted)
    print(f"\n{'→ Sync empfohlen (' + str(needs_sync) + ' Änderungen)' if needs_sync else '✓ Kein Sync nötig — alles aktuell'}\n")
    sys.exit(0 if not needs_sync else 1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="SharePoint incremental sync — lädt neue/geänderte Dateien nach SharePoint hoch."
    )
    parser.add_argument("--config", default="config.yaml", help="Pfad zur config.yaml (default: config.yaml)")
    parser.add_argument("--source", help="Quellordner (überschreibt config.yaml)")
    parser.add_argument("--status", action="store_true", help="Zeigt lokale Änderungen ohne Azure-Verbindung (kein Upload)")
    parser.add_argument("--dry-run", action="store_true", help="Zeigt was hochgeladen würde, ohne tatsächlich hochzuladen")
    parser.add_argument("--force-full-sync", action="store_true", help="Ignoriert State-Datei, lädt alle Dateien erneut hoch")
    parser.add_argument("--log-level", default=None, choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="Log-Level")
    args = parser.parse_args()

    try:
        from sharepoint_sync.config import load_config
        config = load_config(args.config)
    except Exception as e:
        print(f"Konfigurationsfehler: {e}", file=sys.stderr)
        sys.exit(2)

    if args.log_level:
        config.logging.log_level = args.log_level

    from sharepoint_sync.logger import setup_logging
    setup_logging(config.logging.log_dir, config.logging.log_level)

    if args.status:
        _run_status(config, args.source)
        return  # exit handled inside

    if args.source:
        config.source_folder = args.source

    try:
        from sharepoint_sync.auth import SharePointAuthenticator
        from sharepoint_sync.graph_client import GraphClient
        from sharepoint_sync.sharepoint_client import SharePointClient
        from sharepoint_sync.state_manager import StateManager
        from sharepoint_sync.uploader import Uploader

        auth = SharePointAuthenticator(config.tenant_id, config.client_id, config.client_secret)
        graph = GraphClient(auth.get_token, max_retries=config.http.max_retries, backoff_base=config.http.backoff_base_seconds)
        sp_client = SharePointClient(
            graph,
            config.sharepoint.site_url,
            config.sharepoint.target_library,
            config.sharepoint.target_folder,
            config.sync.upload_chunk_size_mb,
        )
        state = StateManager(config.state.state_file)
        uploader = Uploader(config, sp_client, state)

        report = uploader.run(dry_run=args.dry_run, force_full_sync=args.force_full_sync)

        if report.errors:
            log.warning("%d Fehler aufgetreten", len(report.errors))
            sys.exit(2)

    except KeyboardInterrupt:
        log.info("Abgebrochen.")
        sys.exit(0)
    except Exception as e:
        log.exception("Unerwarteter Fehler: %s", e)
        sys.exit(2)
