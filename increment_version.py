#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Bump the ``version`` field in pyproject.toml.

Usage::

    ./increment_version.py patch          # 0.1.0 -> 0.1.1
    ./increment_version.py minor          # 0.1.0 -> 0.2.0
    ./increment_version.py major          # 0.1.0 -> 1.0.0
    ./increment_version.py patch --install  # bump, then reinstall the uv tool
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PYPROJECT = ROOT / "pyproject.toml"

# Anchored to the [project] table's own version key, not any other version= line.
_VERSION_RE = re.compile(
    r'(?ms)^\[project\]\s*$.*?^(version\s*=\s*")(\d+)\.(\d+)\.(\d+)(")'
)


def bump(text: str, part: str) -> tuple[str, str, str]:
    """Return (new_text, old_version, new_version)."""
    match = _VERSION_RE.search(text)
    if match is None:
        raise SystemExit(f"No X.Y.Z version found in [project] of {PYPROJECT}")

    major, minor, patch = (int(match.group(i)) for i in (2, 3, 4))
    old = f"{major}.{minor}.{patch}"

    if part == "major":
        major, minor, patch = major + 1, 0, 0
    elif part == "minor":
        minor, patch = minor + 1, 0
    else:
        patch += 1
    new = f"{major}.{minor}.{patch}"

    start, end = match.span(1)[0], match.span(5)[1]
    new_text = f'{text[:start]}version = "{new}"{text[end:]}'
    return new_text, old, new


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "part",
        nargs="?",
        default="patch",
        choices=("major", "minor", "patch"),
        help="Version component to increment (default: patch)",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="Run 'uv tool install --reinstall .' after bumping",
    )
    parser.add_argument(
        "-n", "--dry-run", action="store_true",
        help="Print the new version without writing pyproject.toml",
    )
    args = parser.parse_args()

    text = PYPROJECT.read_text(encoding="utf-8")
    new_text, old, new = bump(text, args.part)

    if args.dry_run:
        print(f"{old} -> {new} (dry run, nothing written)")
        return

    PYPROJECT.write_text(new_text, encoding="utf-8")
    print(f"{old} -> {new}")

    if args.install:
        result = subprocess.run(
            ["uv", "tool", "install", "--reinstall", "."], cwd=ROOT
        )
        sys.exit(result.returncode)


if __name__ == "__main__":
    main()
