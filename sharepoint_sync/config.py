import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from dotenv import load_dotenv


@dataclass
class SharePointConfig:
    site_url: str
    target_library: str
    target_folder: str


@dataclass
class SyncConfig:
    delete_remote_if_local_deleted: bool = False
    upload_chunk_size_mb: int = 10


@dataclass
class StateConfig:
    state_file: str = "sync_state.json"


@dataclass
class LoggingConfig:
    log_dir: str = "logs"
    log_level: str = "INFO"


@dataclass
class HttpConfig:
    max_retries: int = 5
    backoff_base_seconds: int = 2


@dataclass
class AppConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    source_folder: str
    sharepoint: SharePointConfig
    sync: SyncConfig = field(default_factory=SyncConfig)
    state: StateConfig = field(default_factory=StateConfig)
    logging: LoggingConfig = field(default_factory=LoggingConfig)
    http: HttpConfig = field(default_factory=HttpConfig)


def load_config(config_path: str = "config.yaml") -> AppConfig:
    load_dotenv()

    tenant_id = os.environ.get("AZURE_TENANT_ID", "")
    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")

    missing = [k for k, v in {"AZURE_TENANT_ID": tenant_id, "AZURE_CLIENT_ID": client_id, "AZURE_CLIENT_SECRET": client_secret}.items() if not v]
    if missing:
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    sp_raw = raw.get("sharepoint", {})
    sharepoint = SharePointConfig(
        site_url=sp_raw.get("site_url", ""),
        target_library=sp_raw.get("target_library", "Dokumente"),
        target_folder=sp_raw.get("target_folder", ""),
    )
    if not sharepoint.site_url:
        raise ValueError("config.yaml: sharepoint.site_url is required")

    sync_raw = raw.get("sync", {})
    sync = SyncConfig(
        delete_remote_if_local_deleted=sync_raw.get("delete_remote_if_local_deleted", False),
        upload_chunk_size_mb=sync_raw.get("upload_chunk_size_mb", 10),
    )

    state_raw = raw.get("state", {})
    state = StateConfig(state_file=state_raw.get("state_file", "sync_state.json"))

    log_raw = raw.get("logging", {})
    logging_cfg = LoggingConfig(log_dir=log_raw.get("log_dir", "logs"), log_level=log_raw.get("log_level", "INFO"))

    http_raw = raw.get("http", {})
    http = HttpConfig(max_retries=http_raw.get("max_retries", 5), backoff_base_seconds=http_raw.get("backoff_base_seconds", 2))

    source_folder = raw.get("source_folder", "")
    if not source_folder:
        raise ValueError("config.yaml: source_folder is required")

    return AppConfig(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        source_folder=str(Path(source_folder)),
        sharepoint=sharepoint,
        sync=sync,
        state=state,
        logging=logging_cfg,
        http=http,
    )
