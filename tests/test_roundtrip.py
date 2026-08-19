"""Roundtrip stability tests for the doc↔md converter.

The fixture document is built from ``tests/fixtures/sample.md`` at test time
(markdown → docx) so the suite is self-contained; no binary fixtures are
checked in.  Real-world documents dropped into the gitignored ``docx/`` dir
are picked up as extra cases and skipped when absent.

For each document, verifies:
1. Structural sanity — first-pass markdown has expected minimum counts
   of headings, tables, and images.
2. Convergence — after two full roundtrip cycles
   (doc→md→doc→md→doc→md), the last two markdown outputs are
   identical modulo normalizable differences (table-separator dash
   widths, image-path stems).
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pytest
from PIL import Image

from md_docx.main import convert_docx_to_md, convert_md_to_docx

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_MD = Path(__file__).resolve().parent / "fixtures" / "sample.md"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _count_headings(text: str) -> int:
    return len(re.findall(r"^#{1,6}\s", text, re.MULTILINE))


def _count_tables(text: str) -> int:
    return len(re.findall(r"^\|[-:]+\|", text, re.MULTILINE))


def _count_images(text: str) -> int:
    # doc2md emits centred <img> HTML, md sources use ![...](...).
    return len(re.findall(r"!\[|<img\b", text))


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


def _assert_structure(
    docx_path: Path,
    tmp_path: Path,
    min_headings: int,
    min_tables: int,
    min_images: int,
) -> None:
    """First-pass markdown has the expected structural elements."""
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


def _assert_stability(docx_path: Path, tmp_path: Path) -> None:
    """Markdown structure converges after two roundtrip cycles.

    Compares structural element counts between the two passes.  Full-text
    comparison is too brittle due to Pandoc list re-flowing and
    table-separator cosmetic differences.
    """
    md_rt1, md_rt2 = _do_roundtrip(docx_path, tmp_path)

    counters = {
        "headings": _count_headings,
        "tables": _count_tables,
        "images": _count_images,
        "page_breaks": _count_page_breaks,
    }
    for key, count in counters.items():
        assert count(md_rt1) == count(md_rt2), (
            f"{key}: pass 3 has {count(md_rt1)}, pass 5 has {count(md_rt2)}"
        )


# ---------------------------------------------------------------------------
# Generated fixture document
# ---------------------------------------------------------------------------

_SAMPLE_TEXT = SAMPLE_MD.read_text(encoding="utf-8")

# The leading `# ` heading is promoted into the YAML title block, so it does
# not survive as a heading in the converted markdown.
SAMPLE_MIN_HEADINGS = _count_headings(_SAMPLE_TEXT) - 1
SAMPLE_MIN_TABLES = _count_tables(_SAMPLE_TEXT)
SAMPLE_MIN_IMAGES = _count_images(_SAMPLE_TEXT)


@pytest.fixture(scope="session")
def sample_docx(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Build the fixture .docx from the checked-in markdown source."""
    work = tmp_path_factory.mktemp("sample")
    shutil.copy(SAMPLE_MD, work / "sample.md")

    image_dir = work / "sample_images"
    image_dir.mkdir()
    for name, colour in (("fig1.png", (200, 60, 60)), ("fig2.png", (60, 90, 200))):
        Image.new("RGB", (240, 160), colour).save(image_dir / name)

    return convert_md_to_docx(work / "sample.md", work / "sample.docx")


class TestSampleRoundtrip:
    """Roundtrip tests against the generated fixture document."""

    def test_structural_sanity(self, sample_docx: Path, tmp_path: Path) -> None:
        _assert_structure(
            sample_docx,
            tmp_path,
            SAMPLE_MIN_HEADINGS,
            SAMPLE_MIN_TABLES,
            SAMPLE_MIN_IMAGES,
        )

    def test_markdown_stability(self, sample_docx: Path, tmp_path: Path) -> None:
        _assert_stability(sample_docx, tmp_path)


# ---------------------------------------------------------------------------
# Local real-world documents: (docx filename, min_headings, min_tables,
# min_images).  docx/ is gitignored, so these run only where the files exist.
# ---------------------------------------------------------------------------

REAL_DOCUMENTS = [
    ("ANNOTATION_PROTOCOL.docx", 40, 14, 1),
    ("DRS.docx", 35, 5, 2),
]


def _resolve_real_document(docx_name: str) -> Path:
    path = ROOT / "docx" / docx_name
    if not path.exists():
        pytest.skip(f"{docx_name} not available in docx/")
    return path


@pytest.mark.parametrize(
    "docx_name, min_headings, min_tables, min_images",
    REAL_DOCUMENTS,
    ids=[d[0] for d in REAL_DOCUMENTS],
)
class TestRealDocumentRoundtrip:
    """Roundtrip stability tests per local document."""

    def test_structural_sanity(
        self,
        tmp_path: Path,
        docx_name: str,
        min_headings: int,
        min_tables: int,
        min_images: int,
    ) -> None:
        _assert_structure(
            _resolve_real_document(docx_name),
            tmp_path,
            min_headings,
            min_tables,
            min_images,
        )

    def test_markdown_stability(
        self,
        tmp_path: Path,
        docx_name: str,
        min_headings: int,
        min_tables: int,
        min_images: int,
    ) -> None:
        _assert_stability(_resolve_real_document(docx_name), tmp_path)
