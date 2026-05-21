import logging
import os
from typing import Generator, Tuple

log = logging.getLogger(__name__)

SKIP_FILES = {"Thumbs.db", "desktop.ini", ".DS_Store"}


def _is_office_temp(name: str) -> bool:
    return name.startswith("~$")


class FolderWalker:
    def __init__(self, source_folder: str) -> None:
        self._root = source_folder

    def walk(self) -> Generator[Tuple[str, str], None, None]:
        for dirpath, dirnames, filenames in os.walk(self._root):
            # Skip hidden directories in-place
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]

            for name in filenames:
                if name.startswith(".") or name in SKIP_FILES or _is_office_temp(name):
                    continue
                abs_path = os.path.join(dirpath, name)
                rel = os.path.relpath(abs_path, self._root)
                # Normalize to forward slashes for cross-platform consistency
                rel = rel.replace(os.sep, "/")
                yield abs_path, rel
