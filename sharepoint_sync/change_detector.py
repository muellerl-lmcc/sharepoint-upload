import hashlib
import logging
import os
from enum import Enum, auto
from typing import Optional, Tuple

from sharepoint_sync.state_manager import FileRecord, StateManager

log = logging.getLogger(__name__)

HASH_CHUNK = 8 * 1024 * 1024  # 8 MB read buffer


class FileStatus(Enum):
    NEW = auto()
    MODIFIED = auto()
    UNCHANGED = auto()


def compute_file_fingerprint(path: str) -> Tuple[float, int, str]:
    stat = os.stat(path)
    mtime = stat.st_mtime
    size = stat.st_size
    md5 = _md5(path)
    return mtime, size, md5


def _md5(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        while chunk := f.read(HASH_CHUNK):
            h.update(chunk)
    return h.hexdigest()


class ChangeDetector:
    def __init__(self, state: StateManager) -> None:
        self._state = state

    def classify(self, abs_path: str, relative_path: str) -> Tuple[FileStatus, Optional[Tuple[float, int, str]]]:
        record: Optional[FileRecord] = self._state.get(relative_path)

        if record is None:
            log.debug("NEW: %s", relative_path)
            return FileStatus.NEW, None

        stat = os.stat(abs_path)
        mtime = stat.st_mtime
        size = stat.st_size

        # Fast path: mtime and size unchanged → skip MD5
        if abs(mtime - record.mtime) < 1.0 and size == record.size:
            log.debug("UNCHANGED (fast): %s", relative_path)
            return FileStatus.UNCHANGED, (record.mtime, record.size, record.md5)

        # Slow path: compute MD5
        md5 = _md5(abs_path)
        if md5 == record.md5:
            log.debug("UNCHANGED (md5): %s", relative_path)
            return FileStatus.UNCHANGED, (mtime, size, md5)

        log.debug("MODIFIED: %s", relative_path)
        return FileStatus.MODIFIED, (mtime, size, md5)
