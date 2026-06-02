#!/usr/bin/env python3
"""
Pre-flight analysis for SharePoint sync.
Scans a local folder and reports size, file count, and potential upload issues.

Usage:
    python analyze_folder.py X:\Pfad\zum\Ordner
    python analyze_folder.py X:\Pfad\zum\Ordner --json
"""

import argparse
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Tuple

# SharePoint / Graph API limits
SP_MAX_PATH_LENGTH = 400       # characters, full remote path
SP_MAX_FILENAME_LENGTH = 256   # characters
SP_MAX_FILE_SIZE_GB = 250      # GB, Graph API upload limit
WARN_FILE_SIZE_MB = 100        # warn for files larger than this

# Characters not allowed in SharePoint filenames or path segments
SP_ILLEGAL_CHARS = set('~#%&*{}\\:<>?/|"')
SP_ILLEGAL_NAMES = {
    ".lock", "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}
SP_ILLEGAL_PREFIXES = ("~$",)
SP_ILLEGAL_SUFFIXES = ("_files",)


@dataclass
class Issue:
    severity: str   # ERROR | WARN | INFO
    path: str
    message: str


@dataclass
class AnalysisResult:
    root: str
    total_files: int = 0
    total_dirs: int = 0
    total_size_bytes: int = 0
    skipped_files: int = 0
    issues: List[Issue] = field(default_factory=list)
    largest_files: List[Tuple[int, str]] = field(default_factory=list)  # (size, rel_path)
    extension_stats: dict = field(default_factory=dict)


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"


def _check_name(name: str, rel_path: str, issues: List[Issue]) -> None:
    stem = Path(name).stem.upper()

    if len(name) > SP_MAX_FILENAME_LENGTH:
        issues.append(Issue("ERROR", rel_path, f"Filename too long ({len(name)} chars, max {SP_MAX_FILENAME_LENGTH})"))

    illegal = SP_ILLEGAL_CHARS & set(name)
    if illegal:
        issues.append(Issue("ERROR", rel_path, f"Illegal characters in filename: {' '.join(sorted(illegal))}"))

    if stem in SP_ILLEGAL_NAMES or name.upper() in SP_ILLEGAL_NAMES:
        issues.append(Issue("ERROR", rel_path, f"Reserved filename not allowed in SharePoint: {name}"))

    for prefix in SP_ILLEGAL_PREFIXES:
        if name.startswith(prefix):
            issues.append(Issue("WARN", rel_path, f"Office temp file ('{prefix}' prefix) — will be skipped by sync"))

    for suffix in SP_ILLEGAL_SUFFIXES:
        if Path(name).stem.lower().endswith(suffix):
            issues.append(Issue("WARN", rel_path, f"Filename ending in '{suffix}' may cause issues in SharePoint"))

    if name != name.strip():
        issues.append(Issue("WARN", rel_path, "Filename has leading/trailing whitespace"))

    if name.endswith("."):
        issues.append(Issue("WARN", rel_path, "Filename ends with a dot — not allowed in SharePoint"))


def _check_path_length(rel_path: str, target_folder: str, issues: List[Issue]) -> None:
    full_remote = f"{target_folder}/{rel_path}" if target_folder else rel_path
    if len(full_remote) > SP_MAX_PATH_LENGTH:
        issues.append(Issue("ERROR", rel_path,
            f"Full remote path too long ({len(full_remote)} chars, max {SP_MAX_PATH_LENGTH})"))


def analyze(root: str, target_folder: str = "") -> AnalysisResult:
    result = AnalysisResult(root=root)
    top_n: List[Tuple[int, str]] = []

    SKIP_FILES = {"Thumbs.db", "desktop.ini", ".DS_Store"}

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        result.total_dirs += len(dirnames)

        for name in filenames:
            abs_path = os.path.join(dirpath, name)
            rel_path = os.path.relpath(abs_path, root).replace(os.sep, "/")

            if name.startswith(".") or name in SKIP_FILES:
                result.skipped_files += 1
                continue

            result.total_files += 1

            try:
                size = os.path.getsize(abs_path)
            except OSError as e:
                result.issues.append(Issue("ERROR", rel_path, f"Cannot read file: {e}"))
                continue

            result.total_size_bytes += size

            ext = Path(name).suffix.lower() or "(no extension)"
            result.extension_stats[ext] = result.extension_stats.get(ext, 0) + 1

            # Track largest files (keep top 10)
            top_n.append((size, rel_path))
            top_n.sort(reverse=True)
            top_n = top_n[:10]

            # Size checks
            size_mb = size / (1024 * 1024)
            size_gb = size / (1024 ** 3)
            if size_gb > SP_MAX_FILE_SIZE_GB:
                result.issues.append(Issue("ERROR", rel_path,
                    f"File exceeds Graph API limit ({size_gb:.1f} GB, max {SP_MAX_FILE_SIZE_GB} GB)"))
            elif size_mb > WARN_FILE_SIZE_MB:
                result.issues.append(Issue("WARN", rel_path,
                    f"Large file ({size_mb:.0f} MB) — upload will use chunked transfer"))

            _check_name(name, rel_path, result.issues)
            _check_path_length(rel_path, target_folder, result.issues)

    result.largest_files = top_n
    return result


def print_report(result: AnalysisResult, target_folder: str = "") -> None:
    errors = [i for i in result.issues if i.severity == "ERROR"]
    warnings = [i for i in result.issues if i.severity == "WARN"]

    print("=" * 60)
    print(f"  Folder Analysis: {result.root}")
    print("=" * 60)

    print(f"\n  Dateien:       {result.total_files:>10,}")
    print(f"  Ordner:        {result.total_dirs:>10,}")
    print(f"  Übersprungen:  {result.skipped_files:>10,}  (hidden, Thumbs.db, ...)")
    print(f"  Gesamtgröße:   {_fmt_size(result.total_size_bytes):>10}")

    if result.largest_files:
        print("\n  Größte Dateien:")
        for size, path in result.largest_files:
            print(f"    {_fmt_size(size):>10}  {path}")

    if result.extension_stats:
        print("\n  Dateitypen (Top 10):")
        for ext, count in sorted(result.extension_stats.items(), key=lambda x: -x[1])[:10]:
            print(f"    {count:>6}x  {ext}")

    if target_folder:
        print(f"\n  Zielordner auf SharePoint: {target_folder}")

    print(f"\n  Probleme: {len(errors)} Fehler, {len(warnings)} Warnungen")

    if errors:
        print("\n  [FEHLER] — Upload wird für diese Dateien scheitern:")
        for i in errors:
            print(f"    {i.path}")
            print(f"      → {i.message}")

    if warnings:
        print("\n  [WARNUNG] — Potentielle Probleme:")
        for i in warnings:
            print(f"    {i.path}")
            print(f"      → {i.message}")

    if not errors and not warnings:
        print("\n  Keine Probleme gefunden — Ordner ist sync-bereit.")

    print()
    if errors:
        print("  Ergebnis: NICHT BEREIT — Fehler müssen behoben werden vor dem Sync.")
    elif warnings:
        print("  Ergebnis: BEREIT mit Einschränkungen — Warnungen prüfen.")
    else:
        print("  Ergebnis: BEREIT")
    print("=" * 60)


def main() -> None:
    parser = argparse.ArgumentParser(description="Pre-flight analysis vor SharePoint-Sync")
    parser.add_argument("folder", help="Zu analysierender Ordner (z.B. X:\\Pfad oder /mnt/x/pfad)")
    parser.add_argument("--target-folder", default="", help="Zielordner auf SharePoint (für Pfadlängen-Check)")
    parser.add_argument("--json", action="store_true", help="Ausgabe als JSON")
    args = parser.parse_args()

    root = args.folder
    if not os.path.isdir(root):
        print(f"Fehler: '{root}' ist kein gültiger Ordner.", file=sys.stderr)
        sys.exit(2)

    result = analyze(root, args.target_folder)

    if args.json:
        out = {
            "root": result.root,
            "total_files": result.total_files,
            "total_dirs": result.total_dirs,
            "total_size_bytes": result.total_size_bytes,
            "total_size_human": _fmt_size(result.total_size_bytes),
            "skipped_files": result.skipped_files,
            "largest_files": [{"size": s, "path": p} for s, p in result.largest_files],
            "extension_stats": result.extension_stats,
            "issues": [{"severity": i.severity, "path": i.path, "message": i.message} for i in result.issues],
        }
        print(json.dumps(out, indent=2, ensure_ascii=False))
    else:
        print_report(result, args.target_folder)

    errors = [i for i in result.issues if i.severity == "ERROR"]
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()