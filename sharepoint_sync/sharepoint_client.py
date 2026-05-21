import logging
import math
import os
from urllib.parse import quote

from sharepoint_sync.graph_client import GraphClient, GraphNotFoundError

log = logging.getLogger(__name__)

CHUNK_ALIGN = 327680  # 320 KiB — Graph API requirement
SMALL_FILE_THRESHOLD = 4 * 1024 * 1024  # 4 MB


class SharePointClient:
    def __init__(self, graph: GraphClient, site_url: str, library: str, target_folder: str, chunk_size_mb: int = 10) -> None:
        self._graph = graph
        self._site_url = site_url.rstrip("/")
        self._library = library
        self._target_folder = target_folder.strip("/")
        self._chunk_size = self._align_chunk(chunk_size_mb * 1024 * 1024)
        self._site_id: str | None = None
        self._drive_id: str | None = None
        self._folder_cache: set[str] = set()

    @staticmethod
    def _align_chunk(size: int) -> int:
        return max(CHUNK_ALIGN, (size // CHUNK_ALIGN) * CHUNK_ALIGN)

    def _get_site_id(self) -> str:
        if self._site_id:
            return self._site_id
        # Extract hostname and site path from URL
        # e.g. https://contoso.sharepoint.com/sites/MySite
        url = self._site_url
        if "://" in url:
            url = url.split("://", 1)[1]
        host, *path_parts = url.split("/", 1)
        site_path = path_parts[0] if path_parts else ""
        result = self._graph.get(f"/sites/{host}:/{site_path}")
        self._site_id = result["id"]
        log.debug("Resolved site ID: %s", self._site_id)
        return self._site_id

    def _get_drive_id(self) -> str:
        if self._drive_id:
            return self._drive_id
        site_id = self._get_site_id()
        result = self._graph.get(f"/sites/{site_id}/drives")
        for drive in result.get("value", []):
            if drive.get("name") == self._library:
                self._drive_id = drive["id"]
                log.debug("Resolved drive ID: %s", self._drive_id)
                return self._drive_id
        raise RuntimeError(f"Document library '{self._library}' not found on site")

    def ensure_folder(self, relative_path: str) -> None:
        parts = relative_path.strip("/").split("/")
        if not parts or parts == [""]:
            return

        drive_id = self._get_drive_id()
        current = self._target_folder

        for part in parts:
            path = f"{current}/{part}" if current else part
            if path in self._folder_cache:
                current = path
                continue
            encoded = quote(path, safe="/")
            try:
                self._graph.get(f"/drives/{drive_id}/root:/{encoded}")
                self._folder_cache.add(path)
            except GraphNotFoundError:
                parent_encoded = quote(current, safe="/") if current else ""
                parent_path = f"/drives/{drive_id}/root:/{parent_encoded}:/children" if parent_encoded else f"/drives/{drive_id}/root/children"
                self._graph.post(parent_path, json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
                self._folder_cache.add(path)
                log.debug("Created folder: %s", path)
            current = path

    def upload_file(self, local_path: str, relative_path: str) -> str:
        size = os.path.getsize(local_path)
        if size <= SMALL_FILE_THRESHOLD:
            return self._upload_small(local_path, relative_path)
        return self._upload_large(local_path, relative_path, size)

    def _remote_path(self, relative_path: str) -> str:
        if self._target_folder:
            return f"{self._target_folder}/{relative_path}"
        return relative_path

    def _upload_small(self, local_path: str, relative_path: str) -> str:
        drive_id = self._get_drive_id()
        remote = quote(self._remote_path(relative_path), safe="/")
        with open(local_path, "rb") as f:
            data = f.read()
        headers = {"Content-Type": "application/octet-stream"}
        result = self._graph._request_with_retry(
            "PUT",
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{remote}:/content",
            headers={**self._graph._headers(), **headers},
            data=data,
        )
        log.info("Uploaded (small): %s", relative_path)
        return result["id"]

    def _upload_large(self, local_path: str, relative_path: str, total_size: int) -> str:
        drive_id = self._get_drive_id()
        remote = quote(self._remote_path(relative_path), safe="/")
        session_resp = self._graph.post(
            f"/drives/{drive_id}/root:/{remote}:/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        )
        upload_url = session_resp["uploadUrl"]

        offset = 0
        item_id = None
        with open(local_path, "rb") as f:
            while offset < total_size:
                chunk = f.read(self._chunk_size)
                if not chunk:
                    break
                end = offset + len(chunk) - 1
                content_range = f"bytes {offset}-{end}/{total_size}"
                result = self._graph.put_raw(upload_url, chunk, content_range, len(chunk))
                if result and "id" in result:
                    item_id = result["id"]
                offset += len(chunk)
                log.debug("Uploaded chunk %d-%d of %d for %s", offset - len(chunk), end, total_size, relative_path)

        log.info("Uploaded (large): %s", relative_path)
        return item_id or ""

    def delete_item(self, item_id: str) -> None:
        drive_id = self._get_drive_id()
        self._graph.delete(f"/drives/{drive_id}/items/{item_id}")
        log.info("Deleted remote item: %s", item_id)
