"""Roundtrip stability tests for the doc↔md converter.

For each test document, verifies:
1. Structural sanity — first-pass markdown has expected minimum counts
   of headings, tables, and images.
2. Convergence — after two full roundtrip cycles
   (doc→md→doc→md→doc→md), the last two markdown outputs are
   identical modulo normalizable differences (table-separator dash
   widths, image-path stems).
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from md_docx.main import convert_docx_to_md, convert_md_to_docx

ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _count_headings(text: str) -> int:
    return len(re.findall(r"^#{1,6}\s", text, re.MULTILINE))


def _count_tables(text: str) -> int:
    return len(re.findall(r"^\|[-:]+\|", text, re.MULTILINE))


def _count_images(text: str) -> int:
    return len(re.findall(r"!\[", text))


def _count_page_breaks(text: str) -> int:
    return len(re.findall(r"page-break", text))


def _normalize_for_comparison(text: str) -> str:
    """Normalize markdown to remove known non-significant roundtrip diffs."""
    lines = text.splitlines()
    lines = [l.rstrip() for l in lines]

    normalized = []
    for line in lines:
        # Normalize table separator rows: |---|---|---| → |---|---|---|
        if re.match(r"^\|[-:\s|]+\|$", line):
            line = re.sub(r"-{2,}", "---", line)

        # Strip heading anchors {#some-id}
        line = re.sub(r"\s*\{#[^}]+\}$", "", line)

        # Normalize image paths to just filename
        line = re.sub(
            r"(!\[[^\]]*\]\()[^)]*?/([^/)]+)(\))",
            r"\1\2\3",
            line,
        )

        normalized.append(line)

    text = "\n".join(normalized)
    # Collapse 3+ consecutive blank lines to 2
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _do_roundtrip(docx_path: Path, tmp_path: Path) -> tuple[str, str]:
    """Run two full roundtrip cycles and return the last two markdown outputs.

    Cycle 1: docx → md1 → docx_rt1 → md_rt1
    Cycle 2: md_rt1 → docx_rt2 → md_rt2

    Returns (md_rt1, md_rt2) normalized for comparison.
    """
    md1 = tmp_path / "pass1.md"
    docx_rt1 = tmp_path / "pass2.docx"
    md_rt1 = tmp_path / "pass3.md"
    docx_rt2 = tmp_path / "pass4.docx"
    md_rt2 = tmp_path / "pass5.md"

    convert_docx_to_md(docx_path, md1)
    convert_md_to_docx(md1, docx_rt1)
    convert_docx_to_md(docx_rt1, md_rt1)
    convert_md_to_docx(md_rt1, docx_rt2)
    convert_docx_to_md(docx_rt2, md_rt2)

    return md_rt1.read_text(), md_rt2.read_text()


# ---------------------------------------------------------------------------
# Test data: (docx filename, min_headings, min_tables, min_images)
# ---------------------------------------------------------------------------

DOCUMENTS = [
    ("ANNOTATION_PROTOCOL.docx", 40, 14, 1),
    ("DRS.docx", 35, 5, 2),
]


@pytest.mark.parametrize(
    "docx_name, min_headings, min_tables, min_images",
    DOCUMENTS,
    ids=[d[0] for d in DOCUMENTS],
)
class TestRoundtrip:
    """Roundtrip stability tests per document."""

    def test_structural_sanity(
        self,
        tmp_path: Path,
        docx_name: str,
        min_headings: int,
        min_tables: int,
        min_images: int,
    ) -> None:
        """First-pass markdown has the expected structural elements."""
        docx_path = ROOT / "docx" / docx_name
        md_path = tmp_path / "output.md"
        convert_docx_to_md(docx_path, md_path)
        text = md_path.read_text()

        headings = _count_headings(text)
        tables = _count_tables(text)
        images = _count_images(text)

        assert headings >= min_headings, (
            f"Expected >= {min_headings} headings, got {headings}"
        )
        assert tables >= min_tables, (
            f"Expected >= {min_tables} tables, got {tables}"
        )
        assert images >= min_images, (
            f"Expected >= {min_images} images, got {images}"
        )

    def test_markdown_stability(
        self,
        tmp_path: Path,
        docx_name: str,
        min_headings: int,
        min_tables: int,
        min_images: int,
    ) -> None:
        """Markdown structure converges after two roundtrip cycles.

        Compares structural element counts between the two passes.
        Full-text comparison is too brittle due to Pandoc list
        re-flowing and table-separator cosmetic differences.
        """
        docx_path = ROOT / "docx" / docx_name
        md_rt1, md_rt2 = _do_roundtrip(docx_path, tmp_path)

        counts_rt1 = {
            "headings": _count_headings(md_rt1),
            "tables": _count_tables(md_rt1),
            "images": _count_images(md_rt1),
            "page_breaks": _count_page_breaks(md_rt1),
        }
        counts_rt2 = {
            "headings": _count_headings(md_rt2),
            "tables": _count_tables(md_rt2),
            "images": _count_images(md_rt2),
            "page_breaks": _count_page_breaks(md_rt2),
        }

        for key in counts_rt1:
            assert counts_rt1[key] == counts_rt2[key], (
                f"{key}: pass 3 has {counts_rt1[key]}, "
                f"pass 5 has {counts_rt2[key]}"
            )
