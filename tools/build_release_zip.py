"""Build a clean release ZIP for Excel2Word_with_presets.

Expected repository structure:

Excel2Word_with_presets/
├─ README.md
├─ LICENSE
├─ CHANGELOG.md
├─ requirements.txt
├─ script/
│  └─ excel_rows_to_word_gui.py
├─ presets/
│  ├─ *.json
│  └─ optional helper files, e.g. gdy_2_pola.txt
└─ tools/
   └─ build_release_zip.py

Run from repository root:

    python tools/build_release_zip.py

Output:

    dist/Excel2Word_with_presets_v0.1.0.zip
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

DEFAULT_VERSION = "0.1.0"
PROJECT_NAME = "Excel2Word_with_presets"
MAIN_SCRIPT = Path("script") / "excel_rows_to_word_gui.py"

ROOT_FILES = [
    "README.md",
    "LICENSE",
    "CHANGELOG.md",
    "requirements.txt",
]

PRESET_EXTENSIONS = {".json", ".txt"}


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def should_include_preset(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in PRESET_EXTENSIONS


def collect_files(root: Path) -> list[tuple[Path, Path]]:
    """Return a list of (source_path, archive_relative_path)."""
    files: list[tuple[Path, Path]] = []

    for file_name in ROOT_FILES:
        src = root / file_name
        if src.exists():
            files.append((src, Path(file_name)))
        else:
            print(f"WARNING: missing optional file: {file_name}")

    main_script = root / MAIN_SCRIPT
    if not main_script.exists():
        raise FileNotFoundError(
            f"Missing main script: {MAIN_SCRIPT}. "
            "Move the tested application file to script/excel_rows_to_word_gui.py first."
        )
    files.append((main_script, MAIN_SCRIPT))

    presets_dir = root / "presets"
    if not presets_dir.exists():
        raise FileNotFoundError("Missing presets/ directory.")

    preset_files = sorted(p for p in presets_dir.iterdir() if should_include_preset(p))
    if not preset_files:
        raise FileNotFoundError("No preset files found in presets/.")

    for src in preset_files:
        files.append((src, Path("presets") / src.name))

    return files


def build_zip(version: str, overwrite: bool = False) -> Path:
    root = repo_root()
    dist_dir = root / "dist"
    dist_dir.mkdir(exist_ok=True)

    archive_name = f"{PROJECT_NAME}_v{version}.zip"
    archive_path = dist_dir / archive_name
    top_folder = f"{PROJECT_NAME}_v{version}"

    if archive_path.exists() and not overwrite:
        raise FileExistsError(
            f"Release ZIP already exists: {archive_path}. "
            "Use --overwrite to replace it."
        )

    files = collect_files(root)

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for src, rel in files:
            archive_rel = Path(top_folder) / rel
            zf.write(src, archive_rel.as_posix())

    return archive_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a clean release ZIP for Excel2Word_with_presets."
    )
    parser.add_argument(
        "--version",
        default=DEFAULT_VERSION,
        help=f"Release version without leading v. Default: {DEFAULT_VERSION}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing ZIP file in dist/.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        archive_path = build_zip(version=args.version, overwrite=args.overwrite)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Created release ZIP: {archive_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
