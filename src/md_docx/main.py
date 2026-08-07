#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "pypandoc_binary",
#     "pypandoc",
#     "python-docx>=1.2.0",
#     "Pillow>=10.0.0",
# ]
# ///

"""
Bidirectional converter between ``.docx`` and Markdown.

Subcommands
-----------
**doc2md** — convert ``.docx`` to Markdown::

    ./main.py doc2md <input.docx> [-o output.md]
    ./main.py doc2md -d docx -o md          # directory mirror

**md2doc** — convert Markdown to ``.docx``::

    ./main.py md2doc <input.md> [-o output.docx] [-s reference.docx]

Dependencies
------------
    pypandoc / pypandoc_binary — Pandoc-based conversion engine
    python-docx               — docx post-processing (table keep-together)
    Pillow                    — EMF/WMF blank-image detection
"""

from __future__ import annotations

import argparse
import itertools
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path, PurePosixPath

import pypandoc
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from PIL import Image


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_BRAILLE_FRAMES = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"

# Sentinel injected into docx before Pandoc conversion so page breaks survive.
_PAGE_BREAK_MARKER = "PAGE_BREAK_7f8a9b3c"

# Markdown extension string shared by both directions.
# Disables smart quotes and restricts table format to pipe tables.
_MD_EXTENSIONS = (
    "+pipe_tables"
    "+fenced_code_blocks"
    "+line_blocks"
    "+strikeout"
    "+raw_html"
    "-simple_tables"
    "-multiline_tables"
    "-grid_tables"
    "-smart"
)

_MD_WRITE_FORMAT = "markdown" + _MD_EXTENSIONS
_MD_READ_FORMAT = "markdown+yaml_metadata_block" + _MD_EXTENSIONS

# Lua filter: docx page breaks → HTML <div class="page-break">
_DOC2MD_LUA_FILTER = """\
function RawBlock(el)
    if el.format == "openxml" then
        if el.text:match('w:br w:type="page"') then
            return pandoc.RawBlock("html", '<div class="page-break"></div>')
        end
    end
end

function Para(el)
    if #el.content == 0 then
        return nil
    end
end
"""

# Lua filter: HTML <div class="page-break"> → docx page breaks
_MD2DOC_LUA_FILTER = """\
function Div(el)
    if el.classes:includes("page-break") then
        return pandoc.RawBlock(
            "openxml",
            '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        )
    end
end

function RawBlock(el)
    if el.format == "html" then
        if el.text:match('class="page%-break"') or el.text:match("class='page%-break'") then
            return pandoc.RawBlock(
                "openxml",
                '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
            )
        end
    end
end

function RawInline(el)
    if el.format == "html" and el.text:match("^<img ") then
        local src = el.text:match('src="([^"]+)"')
        if not src then return nil end

        local alt = el.text:match('alt="([^"]*)"') or ""
        local w = el.text:match('width="(%d+)"')
        local h = el.text:match('height="(%d+)"')

        local caption = {}
        if alt ~= "" then
            caption = {pandoc.Str(alt)}
        end

        local attr_pairs = {}
        if w then
            table.insert(attr_pairs, {"width", string.format("%.2fin", tonumber(w) / 96)})
        end
        if h then
            table.insert(attr_pairs, {"height", string.format("%.2fin", tonumber(h) / 96)})
        end

        return pandoc.Image(caption, src, "", pandoc.Attr("", {}, attr_pairs))
    end
end
"""

_YAML_FRONTMATTER = """\
---
title: {title}
---

"""


# ---------------------------------------------------------------------------
# Braille progress indicator
# ---------------------------------------------------------------------------

class _BrailleSpinner:
    """Animated braille spinner displayed while a document is being converted."""

    def __init__(self, label: str) -> None:
        self._label = label
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._animate = sys.stdout.isatty()

    def start(self) -> None:
        """Start the spinner animation in a background thread."""
        if not self._animate:
            return
        self._thread = threading.Thread(target=self._spin, daemon=True)
        self._thread.start()

    def _spin(self) -> None:
        for frame in itertools.cycle(_BRAILLE_FRAMES):
            if self._stop.is_set():
                break
            print(f"\r  {frame} {self._label}", end="", flush=True)
            time.sleep(0.08)

    def stop(self, symbol: str) -> None:
        """Stop the spinner and replace it with *symbol*."""
        self._stop.set()
        if self._thread:
            self._thread.join()
        print(f"\r  {symbol} {self._label}", flush=True)

    def __enter__(self) -> _BrailleSpinner:
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.stop("✗" if exc_type else "✓")


# ---------------------------------------------------------------------------
# Image helpers (EMF/WMF → PNG conversion)
# ---------------------------------------------------------------------------

def _is_blank_image(path: Path, *, threshold: float = 0.01) -> bool:
    """Return True if the image at *path* is blank (single colour or nearly so)."""
    try:
        img = Image.open(str(path)).convert("RGB")
    except Exception:  # noqa: BLE001
        return True

    w, h = img.size
    if w == 0 or h == 0:
        return True

    pixels = img.load()
    total = 0
    non_white = 0
    step = max(1, min(w, h) // 50)
    for x in range(0, w, step):
        for y in range(0, h, step):
            total += 1
            r, g, b = pixels[x, y]
            if (r, g, b) != (255, 255, 255):
                non_white += 1

    return (non_white / total) < threshold if total else True


def _find_soffice() -> str | None:
    """Locate the LibreOffice ``soffice`` binary."""
    path = shutil.which("soffice")
    if path:
        return path
    macos_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if macos_path.exists():
        return str(macos_path)
    return None


def _convert_emf_via_powershell(src: Path, dest: Path) -> bool:
    """Convert EMF/WMF to PNG using PowerShell + .NET (Windows only)."""
    if sys.platform != "win32":
        return False
    ps_script = (
        "Add-Type -AssemblyName System.Drawing; "
        f"$img = [System.Drawing.Image]::FromFile('{src}'); "
        f"$img.Save('{dest}', [System.Drawing.Imaging.ImageFormat]::Png); "
        "$img.Dispose()"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            check=True, capture_output=True, timeout=30,
        )
        return dest.exists() and dest.stat().st_size > 0
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
        return False


def _convert_emf_to_png(src: Path, dest: Path) -> bool:
    """
    Convert an EMF/WMF file to PNG using available system tools.

    Tries (in order): PowerShell/.NET, LibreOffice, ImageMagick, Inkscape,
    Pillow.  Each result is validated — blank images are discarded.
    """

    def _valid(p: Path) -> bool:
        return p.exists() and p.stat().st_size > 0 and not _is_blank_image(p)

    if _convert_emf_via_powershell(src, dest) and _valid(dest):
        return True
    dest.unlink(missing_ok=True)

    soffice = _find_soffice()
    if soffice:
        try:
            subprocess.run(
                [soffice, "--convert-to", "png", "--outdir", str(dest.parent), str(src)],
                check=True, capture_output=True, timeout=60,
            )
            lo_output = dest.parent / f"{src.stem}.png"
            if lo_output.exists() and lo_output != dest:
                lo_output.rename(dest)
            if _valid(dest):
                return True
            dest.unlink(missing_ok=True)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            pass

    if shutil.which("magick"):
        try:
            subprocess.run(
                ["magick", str(src), str(dest)],
                check=True, capture_output=True, timeout=30,
            )
            if _valid(dest):
                return True
            dest.unlink(missing_ok=True)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            pass

    if shutil.which("inkscape"):
        try:
            subprocess.run(
                ["inkscape", str(src), "--export-filename", str(dest)],
                check=True, capture_output=True, timeout=30,
            )
            if _valid(dest):
                return True
            dest.unlink(missing_ok=True)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            pass

    try:
        img = Image.open(str(src))
        img.save(str(dest), "PNG")
        if _valid(dest):
            return True
        dest.unlink(missing_ok=True)
    except Exception:  # noqa: BLE001
        pass

    return False


# ---------------------------------------------------------------------------
# Image cropping (docx a:srcRect → physically cropped media)
# ---------------------------------------------------------------------------

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_CROPPABLE_PART_RE = re.compile(r"word/(document|header\d*|footer\d*)\.xml")

Crop = tuple[float, float, float, float]


def _part_rels(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Return the ``rId`` → target map for an OPC *part*."""
    p = PurePosixPath(part)
    try:
        data = zf.read(str(p.parent / "_rels" / (p.name + ".rels")))
    except KeyError:
        return {}
    root = etree.fromstring(data)
    return {
        rel.get("Id"): rel.get("Target")
        for rel in root.findall(f"{{{_PKG_REL_NS}}}Relationship")
        if rel.get("Id") and rel.get("Target")
    }


def _blip_rel_ids(blip: etree._Element) -> list[str]:
    """Relationship ids referenced by an ``a:blip`` — raster plus any SVG variant."""
    ids = [blip.get(f"{{{_R_NS}}}embed")]
    for el in blip.iter():
        if isinstance(el.tag, str) and etree.QName(el).localname == "svgBlip":
            ids.append(el.get(f"{{{_R_NS}}}embed"))
    return [i for i in ids if i]


def _extract_image_crops(docx_path: Path) -> dict[str, Crop]:
    """
    Map media file name → ``(left, top, right, bottom)`` crop fractions.

    Word crops pictures non-destructively via ``a:srcRect`` while the stored
    media file stays uncropped.  Pandoc extracts the uncropped file but keeps
    the *cropped* display size, which distorts the aspect ratio.

    Only images whose every usage shares a single crop are returned; the same
    media file used with different crops cannot be fixed in place.
    """
    found: dict[str, set[Crop]] = {}

    with zipfile.ZipFile(docx_path) as zf:
        for part in zf.namelist():
            if not _CROPPABLE_PART_RE.fullmatch(part):
                continue
            rels = _part_rels(zf, part)
            root = etree.fromstring(zf.read(part))
            for fill in root.iter():
                if not isinstance(fill.tag, str):
                    continue
                if etree.QName(fill).localname != "blipFill":
                    continue
                src_rect = fill.find(f"{{{_A_NS}}}srcRect")
                blip = fill.find(f"{{{_A_NS}}}blip")
                if src_rect is None or blip is None:
                    continue
                crop = tuple(
                    float(src_rect.get(k, 0)) / 100_000.0 for k in ("l", "t", "r", "b")
                )
                for rid in _blip_rel_ids(blip):
                    target = rels.get(rid)
                    if target:
                        found.setdefault(PurePosixPath(target).name, set()).add(crop)

    result: dict[str, Crop] = {}
    for name, crops in found.items():
        if len(crops) != 1:
            continue
        crop = next(iter(crops))
        if any(crop):
            result[name] = crop
    return result


def _svg_length(value: str | None) -> float | None:
    """Parse the numeric part of an SVG length such as ``"616px"``."""
    if not value:
        return None
    m = re.match(r"\s*(-?\d+(?:\.\d+)?)", value)
    return float(m.group(1)) if m else None


def _crop_svg(path: Path, crop: Crop) -> None:
    """Crop an SVG by narrowing its viewBox (outer ``svg`` clips by default)."""
    left, top, right, bottom = crop
    tree = etree.parse(str(path))
    root = tree.getroot()

    view_box = root.get("viewBox")
    if view_box:
        nums = [float(v) for v in re.split(r"[,\s]+", view_box.strip()) if v]
        if len(nums) != 4:
            return
        x, y, w, h = nums
    else:
        w = _svg_length(root.get("width"))
        h = _svg_length(root.get("height"))
        if not w or not h:
            return
        x = y = 0.0

    new_w, new_h = w * (1 - left - right), h * (1 - top - bottom)
    if new_w <= 0 or new_h <= 0:
        return

    root.set("viewBox", f"{x + w * left:g} {y + h * top:g} {new_w:g} {new_h:g}")
    root.set("width", f"{new_w:g}")
    root.set("height", f"{new_h:g}")
    tree.write(str(path), xml_declaration=True, encoding="utf-8")


def _crop_raster(path: Path, crop: Crop) -> None:
    """Crop a raster image in place, preserving its original format."""
    left, top, right, bottom = crop
    with Image.open(str(path)) as img:
        img.load()
        fmt = img.format
        w, h = img.size
        box = (
            round(w * left),
            round(h * top),
            round(w * (1 - right)),
            round(h * (1 - bottom)),
        )
        new_w, new_h = box[2] - box[0], box[3] - box[1]
        if new_w < 1 or new_h < 1 or new_w > 4 * w or new_h > 4 * h:
            return

        if box[0] < 0 or box[1] < 0 or box[2] > w or box[3] > h:
            # Negative srcRect values pad the picture instead of cropping it.
            mode = "RGB" if fmt in {"JPEG", "BMP"} else "RGBA"
            fill = (255, 255, 255) if mode == "RGB" else (0, 0, 0, 0)
            cropped = Image.new(mode, (new_w, new_h), fill)
            cropped.paste(img.convert(mode), (-box[0], -box[1]))
        else:
            cropped = img.crop(box)
    cropped.save(str(path), format=fmt)


def _apply_image_crops(image_dir: Path, crops: dict[str, Crop]) -> None:
    """Physically crop extracted media so it matches the docx display size."""
    if not crops or not image_dir.is_dir():
        return

    # Match on stem: EMF/WMF sources have already been rewritten to .png.
    by_stem = {PurePosixPath(name).stem: crop for name, crop in crops.items()}

    for path in sorted(image_dir.rglob("*")):
        crop = by_stem.get(path.stem) if path.is_file() else None
        if crop is None:
            continue
        suffix = path.suffix.lower()
        if suffix in {".emf", ".wmf"}:
            continue
        try:
            if suffix == ".svg":
                _crop_svg(path, crop)
            else:
                _crop_raster(path, crop)
        except Exception as exc:  # noqa: BLE001
            print(f"    could not crop {path.name}: {exc}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _straighten_quotes(text: str) -> str:
    """Replace curly/smart quotes with straight ASCII quotes."""
    return (
        text
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2018", "'")
        .replace("\u2019", "'")
    )


# Titles Word templates leave behind, which carry no information about the
# actual document and should fall through to the next candidate.
_PLACEHOLDER_TITLE_RE = re.compile(
    r"^(?:"
    r"untitled|titl?e|titre|document\s*(?:title|name)|doc\s*title|"
    r"insert\s+(?:document\s+)?title(?:\s+here)?|"
    r"[<\[{(].*[>\]})]|"          # <Title>, [Document title], {Title}
    r".*\btemplates?\b.*"
    r")$",
    re.IGNORECASE,
)

# Front-section headings that precede the real content of a report.
_GENERIC_HEADING_RE = re.compile(
    r"^(?:table\s+of\s+contents|contents|toc|index|revision\s+history)$",
    re.IGNORECASE,
)


def _extract_title_from_md(content: str) -> str | None:
    """Extract the document title from the first ``#`` heading or ``%`` title block."""
    for line in content.splitlines():
        stripped = line.strip()
        if stripped.startswith("% "):
            return stripped[2:].strip()
        if stripped.startswith("# "):
            heading = stripped[2:].strip()
            return None if _GENERIC_HEADING_RE.match(heading) else heading
        if stripped and not stripped.startswith("%"):
            break
    return None


def _extract_title_from_docx(docx_path: Path) -> str | None:
    """Extract title from docx core properties metadata."""
    try:
        doc = Document(str(docx_path))
        title = doc.core_properties.title
    except Exception:  # noqa: BLE001
        return None

    title = (title or "").strip()
    if not title or _PLACEHOLDER_TITLE_RE.match(title):
        return None
    return title


# Dangling punctuation Word title pages tend to leave on the title line.
_TITLE_TRAILING_RE = re.compile(r"[\s\u2013\u2014:;,-]+$")


def _format_title_page(md_content: str) -> tuple[str, str | None]:
    """
    Promote a leading title page to an ``#`` heading and return its text.

    Word title pages are plain centred paragraphs, so Pandoc emits them as
    loose body text with no structure.  A title page is recognised as the
    content before the first page break, with no heading in between.
    """
    lines = md_content.split("\n")

    page_break = next(
        (i for i, ln in enumerate(lines) if "page-break" in ln), None,
    )
    heading = next(
        (i for i, ln in enumerate(lines) if ln.startswith("#")), None,
    )
    if page_break is None or (heading is not None and heading < page_break):
        return md_content, None

    head = lines[:page_break]
    title_line = next(
        (
            i for i, ln in enumerate(head)
            if ln.strip() and not ln.lstrip().startswith("<")
        ),
        None,
    )
    if title_line is None:
        return md_content, None

    title = _TITLE_TRAILING_RE.sub("", head[title_line].strip())
    if not title:
        return md_content, None

    head[title_line] = f"# {title}"
    return "\n".join(head + lines[page_break:]), title


def _extract_title_from_frontmatter(md_path: Path) -> str:
    """Extract title from YAML frontmatter, falling back to filename."""
    content = md_path.read_text()
    if not content.startswith("---"):
        return md_path.stem
    end = content.find("---", 3)
    if end == -1:
        return md_path.stem
    frontmatter = content[3:end]
    for line in frontmatter.splitlines():
        if line.startswith("title:"):
            return line.split(":", 1)[1].strip().strip("\"'")
    return md_path.stem


# ---------------------------------------------------------------------------
# Page-break pre-processing (Pandoc 3.x drops w:br page breaks silently)
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _normalize_tables_for_pandoc(doc: Document) -> bool:
    """
    Pre-process tables so Pandoc can emit pipe tables instead of HTML.

    Pandoc falls back to raw HTML tables when it encounters:
    - ``colspan`` / merged cells (``gridSpan > 1``)
    - Multi-paragraph cells (line breaks within a cell)

    This function handles two cases:

    1. **Uniform colspan**: When every row in a table uses the same merge
       pattern (e.g. columns 2+3 always merged), collapse the grid to
       remove the redundant column.
    2. **Multi-paragraph cells**: Merge multiple ``<w:p>`` elements within
       a cell into a single paragraph, joining with `` / ``.

    Returns True if any changes were made.
    """
    ns = _W_NS
    body = doc.element.body
    tables = body.findall(f"{{{ns}}}tbl")
    changed = False

    for tbl in tables:
        grid = tbl.find(f"{{{ns}}}tblGrid")
        if grid is None:
            continue
        grid_cols = grid.findall(f"{{{ns}}}gridCol")
        n_grid = len(grid_cols)
        trs = tbl.findall(f"{{{ns}}}tr")

        if not trs:
            continue

        # --- 1. Detect and fix uniform colspan ---
        # Build per-row merge patterns: list of gridSpan values.
        row_patterns: list[list[int]] = []
        for tr in trs:
            tcs = tr.findall(f"{{{ns}}}tc")
            pattern = []
            for tc in tcs:
                tcPr = tc.find(f"{{{ns}}}tcPr")
                gs = 1
                if tcPr is not None:
                    gs_el = tcPr.find(f"{{{ns}}}gridSpan")
                    if gs_el is not None:
                        gs = int(gs_el.get(f"{{{ns}}}val", "1"))
                pattern.append(gs)
            row_patterns.append(pattern)

        # Check if there's a dominant merge pattern that most rows follow.
        # A "dominant" pattern is one used by > 50% of rows and has spans > 1.
        from collections import Counter
        pattern_counts = Counter(tuple(p) for p in row_patterns)
        dominant_pattern, dominant_count = pattern_counts.most_common(1)[0]

        has_merges = any(s > 1 for s in dominant_pattern)
        is_dominant = dominant_count > len(trs) * 0.5

        if has_merges and is_dominant and sum(dominant_pattern) == n_grid:
            # Normalize outlier rows to match the dominant pattern, then
            # collapse the grid.
            for ri, tr in enumerate(trs):
                current = tuple(row_patterns[ri])
                if current == dominant_pattern:
                    continue
                # Fix this row: merge cells to match dominant pattern.
                tcs = tr.findall(f"{{{ns}}}tc")
                _merge_row_to_pattern(tcs, list(dominant_pattern), ns)
                changed = True

            # Now collapse: remove extra grid columns and gridSpan attrs.
            # Build target column indices (which grid columns survive).
            target_cols: list[int] = []
            gi = 0
            for span in dominant_pattern:
                target_cols.append(gi)
                gi += span

            # Merge grid column widths.
            new_widths: list[int] = []
            gi = 0
            for span in dominant_pattern:
                w = sum(
                    int(grid_cols[gi + j].get(f"{{{ns}}}w", "0"))
                    for j in range(span)
                )
                new_widths.append(w)
                gi += span

            # Remove all grid cols and re-add collapsed ones.
            for gc in grid_cols:
                grid.remove(gc)
            for w in new_widths:
                new_gc = OxmlElement("w:gridCol")
                new_gc.set(f"{{{ns}}}w", str(w))
                grid.append(new_gc)

            # Remove gridSpan from all cells and update cell widths.
            for tr in trs:
                tcs = tr.findall(f"{{{ns}}}tc")
                ci = 0
                for tc in tcs:
                    tcPr = tc.find(f"{{{ns}}}tcPr")
                    if tcPr is not None:
                        gs_el = tcPr.find(f"{{{ns}}}gridSpan")
                        if gs_el is not None:
                            tcPr.remove(gs_el)
                        # Update cell width to match new grid.
                        tcW = tcPr.find(f"{{{ns}}}tcW")
                        if tcW is not None and ci < len(new_widths):
                            tcW.set(f"{{{ns}}}w", str(new_widths[ci]))
                    ci += 1

            changed = True

        # --- 2. Merge multi-paragraph cells ---
        for tr in trs:
            tcs = tr.findall(f"{{{ns}}}tc")
            for tc in tcs:
                paras = tc.findall(f"{{{ns}}}p")
                if len(paras) <= 1:
                    continue
                # Collect text from all paragraphs.
                all_texts: list[str] = []
                for p in paras:
                    runs = p.findall(f".//{{{ns}}}t")
                    text = "".join(r.text or "" for r in runs).strip()
                    if text:
                        all_texts.append(text)

                if not all_texts:
                    continue

                # Keep first paragraph, remove the rest.
                merged_text = " / ".join(all_texts)
                first_p = paras[0]
                # Clear existing runs in first paragraph.
                for r in first_p.findall(f"{{{ns}}}r"):
                    first_p.remove(r)
                # Add merged text as single run.
                new_r = OxmlElement("w:r")
                new_t = OxmlElement("w:t")
                new_t.text = merged_text
                new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                new_r.append(new_t)
                first_p.append(new_r)
                # Remove extra paragraphs.
                for p in paras[1:]:
                    tc.remove(p)
                changed = True

    return changed


def _merge_row_to_pattern(
    tcs: list[etree._Element],
    target_pattern: list[int],
    ns: str,
) -> None:
    """
    Adjust cells in a table row to match *target_pattern* merge pattern.

    When a row has fewer merges than the dominant pattern (e.g. 3 cells
    where the pattern expects 2 cells with one spanning 2 grid columns),
    merge the appropriate cells.
    """
    # Current spans for this row.
    current_spans: list[int] = []
    for tc in tcs:
        tcPr = tc.find(f"{{{ns}}}tcPr")
        gs = 1
        if tcPr is not None:
            gs_el = tcPr.find(f"{{{ns}}}gridSpan")
            if gs_el is not None:
                gs = int(gs_el.get(f"{{{ns}}}val", "1"))
        current_spans.append(gs)

    if len(current_spans) == len(target_pattern):
        # Same number of cells — just fix gridSpan values.
        for tc, target_gs in zip(tcs, target_pattern):
            tcPr = tc.find(f"{{{ns}}}tcPr")
            if target_gs > 1:
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    tc.insert(0, tcPr)
                gs_el = tcPr.find(f"{{{ns}}}gridSpan")
                if gs_el is None:
                    gs_el = OxmlElement("w:gridSpan")
                    tcPr.append(gs_el)
                gs_el.set(f"{{{ns}}}val", str(target_gs))
        return

    # Different number of cells — need to merge/remove some.
    # Walk through grid positions to align current cells with target.
    tr = tcs[0].getparent()
    cur_gi = 0  # current grid index
    cur_ci = 0  # current cell index
    tar_ci = 0  # target pattern index

    while tar_ci < len(target_pattern) and cur_ci < len(tcs):
        target_gs = target_pattern[tar_ci]
        current_gs = current_spans[cur_ci] if cur_ci < len(current_spans) else 1

        if current_gs == target_gs:
            cur_gi += current_gs
            cur_ci += 1
            tar_ci += 1
            continue

        if target_gs > current_gs:
            # Need to absorb subsequent cells into this one.
            cells_to_merge = target_gs - current_gs
            primary = tcs[cur_ci]
            for _ in range(cells_to_merge):
                next_ci = cur_ci + 1
                if next_ci < len(tcs):
                    # Append text from the cell being absorbed.
                    absorbed = tcs[next_ci]
                    texts = absorbed.findall(f".//{{{ns}}}t")
                    absorbed_text = "".join(
                        t.text or "" for t in texts
                    ).strip()
                    if absorbed_text:
                        # Append to primary cell.
                        last_p = primary.findall(f"{{{ns}}}p")[-1]
                        new_r = OxmlElement("w:r")
                        new_t = OxmlElement("w:t")
                        new_t.text = " " + absorbed_text
                        new_t.set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                        new_r.append(new_t)
                        last_p.append(new_r)
                    tr.remove(absorbed)
                    tcs.pop(next_ci)
                    current_spans.pop(next_ci)

            # Set gridSpan on the primary cell.
            tcPr = primary.find(f"{{{ns}}}tcPr")
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                primary.insert(0, tcPr)
            gs_el = tcPr.find(f"{{{ns}}}gridSpan")
            if gs_el is None:
                gs_el = OxmlElement("w:gridSpan")
                tcPr.append(gs_el)
            gs_el.set(f"{{{ns}}}val", str(target_gs))

            cur_gi += target_gs
            cur_ci += 1
            tar_ci += 1
        else:
            # Target wants fewer span — skip for now.
            cur_gi += current_gs
            cur_ci += 1
            tar_ci += 1


def _flatten_revised_fields(body: etree._Element) -> bool:
    """
    Hoist runs out of ``w:ins`` wrappers that contain Word field characters.

    A ``fldChar`` nested inside a tracked insertion breaks Pandoc's field
    parser: the field never closes, and every following paragraph is swallowed
    into it — turning later lists into empty bullets and merging their text
    into the next heading.  Unwrapping accepts the insertion, which is what
    Pandoc's default ``--track-changes=accept`` does anyway.
    """
    changed = False
    for ins in list(body.iter(f"{{{_W_NS}}}ins")):
        if ins.find(f".//{{{_W_NS}}}fldChar") is None:
            continue
        parent = ins.getparent()
        index = list(parent).index(ins)
        for child in reversed(list(ins)):
            parent.insert(index, child)
        parent.remove(ins)
        changed = True
    return changed


def _preprocess_docx(docx_path: Path) -> Path:
    """
    Create a temp copy of *docx_path* with pre-processing applied:

    1. Replace explicit page breaks with marker paragraphs (Pandoc 3.x
       silently discards ``<w:br w:type="page"/>``).
    2. Normalize tables so Pandoc can emit pipe tables instead of HTML.
    3. Unwrap tracked insertions around Word field characters.

    Returns the original path if no changes were needed.
    """
    doc = Document(str(docx_path))
    body = doc.element.body
    changed = False

    # --- Page-break markers ---
    breaks: list[etree._Element] = body.findall(
        f".//{{{_W_NS}}}br[@{{{_W_NS}}}type='page']",
    )
    if breaks:
        changed = True
        for br in breaks:
            run = br.getparent()   # w:r
            para = run.getparent()  # w:p

            marker_p = OxmlElement("w:p")
            marker_r = OxmlElement("w:r")
            marker_t = OxmlElement("w:t")
            marker_t.text = _PAGE_BREAK_MARKER
            marker_r.append(marker_t)
            marker_p.append(marker_r)

            para.addnext(marker_p)
            run.remove(br)

    # --- Table normalization ---
    if _normalize_tables_for_pandoc(doc):
        changed = True

    if _flatten_revised_fields(body):
        changed = True

    if not changed:
        return docx_path

    tmp = Path(tempfile.mktemp(suffix=".docx"))
    doc.save(str(tmp))
    return tmp


# ---------------------------------------------------------------------------
# Table column width preservation
# ---------------------------------------------------------------------------

# Minimum total separator width to ensure Pandoc uses proportional sizing
# (must exceed --columns default of 72).
_MIN_SEPARATOR_WIDTH = 80

_PIPE_TABLE_SEP_RE = re.compile(
    r"^\|?"           # optional leading pipe
    r"(\s*:?-+:?\s*\|)+"  # one or more dash cells
    r"\s*:?-+:?\s*"   # final cell (no trailing pipe required)
    r"\|?\s*$",        # optional trailing pipe
)


def _extract_table_col_widths(docx_path: Path) -> list[list[float]]:
    """
    Extract column width proportions for every table in *docx_path*.

    Returns a list of tables, each being a list of floats summing to 1.0.
    """
    doc = Document(str(docx_path))
    result: list[list[float]] = []

    for table in doc.tables:
        tbl = table._tbl
        grid = tbl.find(f".//{{{_W_NS}}}tblGrid")
        if grid is None:
            result.append([])
            continue
        cols = grid.findall(f"{{{_W_NS}}}gridCol")
        widths = [int(c.get(f"{{{_W_NS}}}w", "0")) for c in cols]
        total = sum(widths)
        if total == 0:
            result.append([])
        else:
            result.append([w / total for w in widths])

    return result


def _clamp_proportions(proportions: list[float]) -> list[float]:
    """
    Apply a minimum column width floor to prevent extremely narrow columns.

    Prettier-formatted tables set separator dashes proportional to the
    longest cell content, which can produce extreme skew (e.g. 7%:93%
    for a short label column next to a long description).  This function
    ensures every column gets at least ``1 / (2 * n_cols)`` of the total
    width, redistributing the excess from wider columns proportionally.
    """
    n = len(proportions)
    if n <= 1:
        return proportions

    floor = 1.0 / (2 * n)
    clamped = list(proportions)
    deficit = 0.0

    # First pass: lift columns below the floor.
    for i, p in enumerate(clamped):
        if p < floor:
            deficit += floor - p
            clamped[i] = floor

    if deficit <= 0:
        return clamped

    # Second pass: reduce columns above the floor proportionally.
    above = [(i, clamped[i]) for i in range(n) if clamped[i] > floor]
    total_above = sum(v for _, v in above)
    if total_above <= 0:
        return clamped
    for i, v in above:
        clamped[i] = v - deficit * (v / total_above)

    return clamped


def _extract_md_table_proportions(md_content: str) -> list[list[float]]:
    """
    Extract column width proportions from pipe-table separator lines.

    Returns a list of tables, each being a list of floats summing to ~1.0,
    derived from the relative dash counts in each separator line.  A
    minimum-width floor is applied to prevent extremely narrow columns
    caused by Prettier-style dash equalization.
    """
    result: list[list[float]] = []
    for line in md_content.split("\n"):
        if not _PIPE_TABLE_SEP_RE.match(line):
            continue
        cells = [c for c in line.split("|") if c.strip()]
        dash_counts = []
        for cell in cells:
            stripped = cell.strip()
            dashes = sum(1 for ch in stripped if ch == "-")
            dash_counts.append(max(dashes, 1))
        total = sum(dash_counts)
        if total > 0:
            props = [d / total for d in dash_counts]
            result.append(_clamp_proportions(props))
        else:
            result.append([])
    return result


def _adjust_table_separators(md_content: str, col_widths: list[list[float]]) -> str:
    """
    Rewrite pipe-table separator lines so that dash counts reflect
    the original column width proportions from the docx.

    This ensures Pandoc uses the correct proportions when converting
    back to docx.
    """
    if not col_widths:
        return md_content

    lines = md_content.split("\n")
    table_idx = 0
    i = 0

    while i < len(lines):
        line = lines[i]
        if not _PIPE_TABLE_SEP_RE.match(line):
            i += 1
            continue

        # Found a separator line — determine column count from it.
        # Split by | but ignore leading/trailing empties.
        cells = [c for c in line.split("|") if c.strip()]
        ncols = len(cells)

        # Find matching table widths (same column count).
        widths: list[float] = []
        while table_idx < len(col_widths):
            candidate = col_widths[table_idx]
            table_idx += 1
            if len(candidate) == ncols:
                widths = candidate
                break

        if not widths:
            i += 1
            continue

        # Rebuild separator line with proportional dashes.
        # Preserve alignment markers (:) from original cells.
        alignments: list[tuple[bool, bool]] = []
        for cell in cells:
            stripped = cell.strip()
            left = stripped.startswith(":")
            right = stripped.endswith(":")
            alignments.append((left, right))

        # Calculate dash counts (minimum 3 per column).
        total_dashes = max(_MIN_SEPARATOR_WIDTH, ncols * 3)
        dash_counts = [max(3, round(w * total_dashes)) for w in widths]
        # Adjust rounding to hit the target total.
        diff = total_dashes - sum(dash_counts)
        if diff != 0:
            # Add/subtract from the widest column.
            widest = dash_counts.index(max(dash_counts))
            dash_counts[widest] += diff

        # Build new separator cells.
        new_cells: list[str] = []
        for (left, right), count in zip(alignments, dash_counts):
            dashes = "-" * count
            if left and right:
                cell_str = f":{dashes[:-2]}:"
            elif left:
                cell_str = f":{dashes[:-1]}"
            elif right:
                cell_str = f"{dashes[:-1]}:"
            else:
                cell_str = dashes
            new_cells.append(cell_str)

        lines[i] = "|" + "|".join(new_cells) + "|"
        i += 1

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Image group post-processing (doc2md cleanup)
# ---------------------------------------------------------------------------

# Pattern matching a Pandoc image: ![alt](src){width="..." height="..."}
_IMG_RE = re.compile(
    r'!\[(?P<alt>[^\]]*)\]'           # ![alt]
    r'\((?P<src>[^)]+)\)'             # (src)
    r'(?:\{(?P<attrs>[^}]*)\})?'      # optional {width="..." height="..."}
)

# Pattern matching a width or height attribute value with many decimal places
_DIM_RE = re.compile(
    r'((?:width|height)=")'           # attribute name + opening quote
    r'(\d+\.\d{3,})'                  # number with 3+ decimal places
    r'(in")',                          # unit + closing quote
)

# Pattern matching a figure caption line (with optional blockquote prefix)
_FIGURE_CAPTION_RE = re.compile(
    r'^(?:>\s*)?(?:Figure\s+\d+)', re.IGNORECASE,
)


def _round_image_dimensions(line: str) -> str:
    """Round width/height values to 2 decimal places for readability."""
    def _round_match(m: re.Match) -> str:
        prefix, value, suffix = m.group(1), m.group(2), m.group(3)
        rounded = f"{float(value):.2f}"
        return f"{prefix}{rounded}{suffix}"
    return _DIM_RE.sub(_round_match, line)


def _post_process_images(md_content: str) -> str:
    """
    Clean up image groups produced by Pandoc's docx-to-markdown conversion.

    Improvements:
    - Strip blockquote markers (``>``) from lines containing only images
      and their immediately following figure captions.  Pandoc adds these
      because the source paragraphs had indentation, but the indentation
      is lost in the roundtrip anyway.
    - Round ``width``/``height`` attribute values to 2 decimal places.
    - Ensure a space between adjacent inline images on the same line
      for readability (``}![`` → ``} ![``).
    """
    lines = md_content.split("\n")
    result: list[str] = []
    i = 0

    def _is_bare_blockquote(ln: str) -> bool:
        """True if the line is just ``>`` with optional whitespace."""
        return ln.strip() == ">" or ln.strip() == ""

    def _is_image_only(ln: str) -> bool:
        """True if the line contains only image references (ignoring ``>`` prefix)."""
        stripped = ln.lstrip("> ").strip()
        if not stripped or not _IMG_RE.search(stripped):
            return False
        remaining = _IMG_RE.sub("", stripped).strip()
        return not remaining

    def _clean_image_line(ln: str) -> str:
        """Strip blockquote prefix, round dimensions, ensure spacing."""
        clean = ln.lstrip("> ") if ln.startswith(">") else ln
        clean = _round_image_dimensions(clean)
        clean = re.sub(r'\}\s*!\[', '} ![', clean)
        return clean

    def _clean_caption(ln: str) -> str:
        """Strip blockquote prefix from a caption line."""
        if ln.startswith(">"):
            return ln.lstrip("> ").strip()
        return ln

    while i < len(lines):
        line = lines[i]

        if not _is_image_only(line):
            # Non-image line — just round any inline image dimensions.
            result.append(_round_image_dimensions(line))
            i += 1
            continue

        # --- Image-only line found ---
        result.append(_clean_image_line(line))
        i += 1

        # Try to consume: optional blank/`>` separator → caption line.
        consumed_blank = False
        if i < len(lines) and _is_bare_blockquote(lines[i]):
            consumed_blank = True
            i += 1

        if i < len(lines) and _FIGURE_CAPTION_RE.match(lines[i]):
            result.append("")
            result.append(_clean_caption(lines[i]))
            i += 1

            # Consume any trailing bare `>` lines (blockquote continuation
            # separators between consecutive image+caption pairs).
            consumed_separator = False
            while i < len(lines) and lines[i].strip() == ">":
                consumed_separator = True
                i += 1
            # If we consumed separators and the next line is still content
            # (not already a blank line), add a blank line to preserve
            # paragraph breaks between figure groups.
            if consumed_separator:
                result.append("")
        elif consumed_blank:
            # We consumed a blank but it wasn't followed by a caption —
            # emit an empty line to preserve the paragraph break.
            result.append("")

    return "\n".join(result)


# Regex for images that carry Pandoc-style width/height attributes.
_IMG_WITH_DIMS_RE = re.compile(
    r'!\[([^\]]*)\]'                         # ![alt]
    r'\(([^)]+)\)'                            # (src)
    r'\{([^}]*(?:width|height)[^}]*)\}',      # {attrs containing width/height}
)

# Conversion factor: standard CSS/screen resolution.
_DPI = 96


def _md_images_to_html(md_content: str) -> str:
    """
    Convert Pandoc ``![alt](src){width="…" height="…"}`` image syntax to
    HTML ``<img>`` tags with pixel dimensions.

    Standard markdown renderers (VS Code, GitHub) do not understand Pandoc's
    curly-brace attribute syntax and display it as raw text.  HTML ``<img>``
    tags with ``width``/``height`` in pixels render correctly everywhere.

    The companion Lua filter in ``_MD2DOC_LUA_FILTER`` converts these tags
    back to Pandoc ``Image`` elements with inch-based dimensions during
    ``md2doc`` conversion.
    """

    def _replace(m: re.Match) -> str:
        alt = (m.group(1) or "").replace('"', '&quot;')
        src = m.group(2).replace('"', '&quot;')
        attrs = m.group(3)

        parts = [f'<img src="{src}"']
        if alt:
            parts.append(f'alt="{alt}"')

        w_m = re.search(r'width="([^"]+)"', attrs)
        h_m = re.search(r'height="([^"]+)"', attrs)

        if w_m:
            val = w_m.group(1)
            if val.endswith("in"):
                px = round(float(val[:-2]) * _DPI)
                parts.append(f'width="{px}"')
        if h_m:
            val = h_m.group(1)
            if val.endswith("in"):
                px = round(float(val[:-2]) * _DPI)
                parts.append(f'height="{px}"')

        return " ".join(parts) + ">"

    return _IMG_WITH_DIMS_RE.sub(_replace, md_content)


_IMG_TAG_RE = re.compile(r"<img\s[^>]*>")


def _center_image_lines(md_content: str) -> str:
    """
    Wrap image-only lines in ``<p align="center">``.

    The deprecated ``align`` attribute is used rather than inline CSS because
    GitHub's HTML sanitizer strips ``style`` but keeps ``align``.
    """
    lines = md_content.split("\n")
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped.startswith("<img"):
            continue
        if _IMG_TAG_RE.sub("", stripped).strip():
            continue
        lines[i] = f'<p align="center">{stripped}</p>'
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Table of contents
# ---------------------------------------------------------------------------

# Word TOC field converted by Pandoc: [Title [page](#anchor)](#anchor)
_TOC_ENTRY_RE = re.compile(
    r'^\[(?P<text>.+?)\s*\[\d+\]\([^)]*\)\]\((?P<anchor>[^)]*)\)$'
)

_SECTION_NUM_RE = re.compile(r'^(\d+(?:\.\d+)*)\s')


def _toc_entry(match: re.Match) -> str | None:
    """Render one matched TOC entry as an indented list item."""
    text = match.group("text").strip()
    anchor = match.group("anchor").strip()

    num = _SECTION_NUM_RE.match(text)
    indent = "  " * (num.group(1).count(".") if num else 0)

    # `#_TocNNNNN` is an unresolved Word bookmark — it links nowhere.
    if not anchor.startswith("#") or anchor.startswith("#_Toc"):
        # Unnumbered and unlinkable means it is the TOC's own self-reference.
        return f"{indent}- {text}" if num else None
    return f"{indent}- [{text}]({anchor})"


def _fix_toc(md_content: str) -> str:
    """
    Rewrite Pandoc's table of contents into a nested bullet list.

    Word TOC fields become ``[Title [page](#anchor)](#anchor)`` — a link nested
    inside a link, which no Markdown renderer supports, so entries show up as
    literal text rather than navigation.  Page numbers are dropped and the
    heading numbering is used to indent entries by level.
    """
    lines = md_content.split("\n")
    result: list[str] = []
    i = 0

    while i < len(lines):
        if not _TOC_ENTRY_RE.match(lines[i].strip()):
            result.append(lines[i])
            i += 1
            continue

        blanks = 0
        while i < len(lines):
            stripped = lines[i].strip()
            if not stripped:
                blanks += 1
                i += 1
                continue
            match = _TOC_ENTRY_RE.match(stripped)
            if match is None:
                break
            entry = _toc_entry(match)
            if entry is not None:
                result.append(entry)
            blanks = 0
            i += 1
        i -= blanks

    return "\n".join(result)


# ---------------------------------------------------------------------------
# EMF post-processing
# ---------------------------------------------------------------------------

def _post_process_emf_images(image_dir: Path, md_path: Path) -> None:
    """Convert any EMF/WMF files in *image_dir* to PNG and fix markdown refs."""
    if not image_dir.exists():
        return
    emf_files = list(image_dir.rglob("*.emf")) + list(image_dir.rglob("*.wmf"))
    if not emf_files:
        return

    md_content = md_path.read_text(encoding="utf-8")
    changed = False
    for emf_file in emf_files:
        png_dest = emf_file.with_suffix(".png")
        if _convert_emf_to_png(emf_file, png_dest):
            rel_emf = str(emf_file.relative_to(md_path.parent))
            rel_png = str(png_dest.relative_to(md_path.parent))
            md_content = md_content.replace(rel_emf, rel_png)
            emf_file.unlink()
            changed = True

    if changed:
        md_path.write_text(md_content, encoding="utf-8")


# ---------------------------------------------------------------------------
# doc2md conversion (pypandoc-based)
# ---------------------------------------------------------------------------

def convert_docx_to_md(
    docx_path: str | Path,
    md_path: str | Path | None = None,
) -> Path:
    """Convert a .docx file to Markdown with extracted images."""
    docx_path = Path(docx_path)
    if md_path is None:
        md_path = docx_path.with_suffix(".md")
    else:
        md_path = Path(md_path)

    md_path.parent.mkdir(parents=True, exist_ok=True)

    image_dir_name = f"{md_path.stem}_images"
    image_dir = md_path.parent / image_dir_name

    # Extract title from docx metadata before conversion
    docx_title = _extract_title_from_docx(docx_path)

    # Extract table column widths from original docx for later post-processing.
    table_col_widths = _extract_table_col_widths(docx_path)

    # Word crops pictures non-destructively; record crops so the extracted
    # media can be cropped for real (otherwise it renders squeezed).
    image_crops = _extract_image_crops(docx_path)

    # Pre-process: inject page-break markers, normalize tables
    preprocessed = _preprocess_docx(docx_path)
    is_temp = preprocessed != docx_path

    # Write Lua filter to temp file
    lua_path = Path(tempfile.mktemp(suffix=".lua"))
    lua_path.write_text(_DOC2MD_LUA_FILTER)

    try:
        pypandoc.convert_file(
            str(preprocessed),
            "md",
            outputfile=str(md_path),
            extra_args=[
                "--lua-filter", str(lua_path),
                "--wrap=none",
                "--to", _MD_WRITE_FORMAT,
                "--extract-media", str(md_path.parent),
            ],
        )
    finally:
        lua_path.unlink(missing_ok=True)
        if is_temp:
            preprocessed.unlink(missing_ok=True)

    # Rename Pandoc's default media/ dir to <stem>_images/
    media_dir = md_path.parent / "media"
    md_content = md_path.read_text(encoding="utf-8")

    if media_dir.exists() and any(media_dir.iterdir()):
        if image_dir.exists():
            shutil.rmtree(image_dir)
        media_dir.rename(image_dir)

        # Rewrite image references only inside link targets and src attributes.
        # A blanket text replace would also hit prose such as "median".
        prefixes = {str(md_path.parent / "media"), "media"}
        media_ref_re = re.compile(
            r'(?P<lead>\]\(|src=["\'])(?P<dots>\./)?'
            r"(?:" + "|".join(
                re.escape(p) for p in sorted(prefixes, key=len, reverse=True)
            ) + r")/",
        )
        md_content = media_ref_re.sub(
            lambda m: f"{m['lead']}{m['dots'] or ''}{image_dir_name}/",
            md_content,
        )
    elif media_dir.exists():
        media_dir.rmdir()

    # Straighten smart quotes
    md_content = _straighten_quotes(md_content)

    # Replace page-break markers injected during pre-processing
    md_content = md_content.replace(
        _PAGE_BREAK_MARKER, '<div class="page-break"></div>',
    )

    md_content, page_title = _format_title_page(md_content)

    # Adjust pipe-table separator dashes to preserve original column proportions
    md_content = _adjust_table_separators(md_content, table_col_widths)

    md_content = _fix_toc(md_content)

    # Clean up image groups (strip blockquotes, round dimensions, spacing)
    md_content = _post_process_images(md_content)

    # Convert Pandoc image attributes to HTML <img> tags for clean rendering
    md_content = _md_images_to_html(md_content)

    md_content = _center_image_lines(md_content)

    # Extract title: prefer the title page, then docx metadata, then filename
    title = (
        page_title
        or docx_title
        or _extract_title_from_md(md_content)
        or md_path.stem
    )
    final_content = _YAML_FRONTMATTER.format(title=title) + md_content

    md_path.write_text(final_content, encoding="utf-8")

    # Post-process EMF/WMF images
    _post_process_emf_images(image_dir, md_path)

    _apply_image_crops(image_dir, image_crops)

    # Clean up empty image dir
    if image_dir.exists() and not any(image_dir.iterdir()):
        image_dir.rmdir()

    return md_path


# ---------------------------------------------------------------------------
# md2doc conversion (pypandoc-based)
# ---------------------------------------------------------------------------

def _keep_tables_together(docx_path: Path) -> None:
    """Post-process docx to prevent tables from breaking across pages."""
    doc = Document(str(docx_path))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    ppr = paragraph._p.get_or_add_pPr()
                    keep_next = OxmlElement("w:keepNext")
                    keep_next.set(qn("w:val"), "true")
                    ppr.append(keep_next)
                    keep_lines = OxmlElement("w:keepLines")
                    keep_lines.set(qn("w:val"), "true")
                    ppr.append(keep_lines)
    doc.save(str(docx_path))


def _center_image_paragraphs(docx_path: Path) -> None:
    """Center paragraphs whose only content is one or more images."""
    doc = Document(str(docx_path))
    changed = False
    for para in doc.paragraphs:
        if para.text.strip() or not para._p.findall(f".//{qn('w:drawing')}"):
            continue
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        changed = True
    if changed:
        doc.save(str(docx_path))


def _apply_table_col_widths(docx_path: Path, col_widths: list[list[float]]) -> None:
    """
    Post-process *docx_path* to set table column widths to match proportions.

    Uses the proportions from *col_widths* (one entry per table, each a list
    of floats summing to 1.0).  Tables are aligned with output tables by
    column count rather than strict index, so an extra or missing table in
    either list doesn't throw off all subsequent matches.
    """
    if not col_widths:
        return

    doc = Document(str(docx_path))
    ns = _W_NS

    # Build list of (table, n_cols) for the output document.
    out_tables: list[tuple] = []
    for table in doc.tables:
        grid = table._tbl.find(f".//{{{ns}}}tblGrid")
        n = len(grid.findall(f"{{{ns}}}gridCol")) if grid is not None else 0
        out_tables.append((table, n))

    # Align output tables to reference widths with bounded look-ahead.
    # Allow skipping at most 2 reference entries to handle a missing table
    # in the markdown.  If no match is found within the window, advance
    # both pointers (the table structures diverged).
    _MAX_SKIP = 2
    ref_idx = 0
    for table, n_cols in out_tables:
        if ref_idx >= len(col_widths):
            break

        # Find a matching reference within the look-ahead window.
        match_offset = None
        for offset in range(_MAX_SKIP + 1):
            candidate = ref_idx + offset
            if candidate >= len(col_widths):
                break
            if len(col_widths[candidate]) == n_cols:
                match_offset = offset
                break

        if match_offset is None:
            # No match in window — treat output table as extra; don't
            # consume a reference entry.
            continue

        ref_idx += match_offset
        widths = col_widths[ref_idx]
        ref_idx += 1
        tbl = table._tbl

        grid = tbl.find(f".//{{{ns}}}tblGrid")
        if grid is None:
            continue
        grid_cols = grid.findall(f"{{{ns}}}gridCol")

        current_widths = [int(c.get(f"{{{ns}}}w", "0")) for c in grid_cols]
        total = sum(current_widths) or 9406  # fallback: ~6.5in in twips

        # Set gridCol widths according to proportions.
        for col_el, proportion in zip(grid_cols, widths):
            col_el.set(f"{{{ns}}}w", str(round(proportion * total)))

        # Also update cell widths (tcW) in each row to match.
        for row in table.rows:
            cells = row.cells
            if len(cells) != len(widths):
                continue
            for cell, proportion in zip(cells, widths):
                tc = cell._tc
                tcPr = tc.find(f"{{{ns}}}tcPr")
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    tc.insert(0, tcPr)
                tcW = tcPr.find(f"{{{ns}}}tcW")
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(f"{{{ns}}}w", str(round(proportion * total)))
                tcW.set(f"{{{ns}}}type", "dxa")

    doc.save(str(docx_path))


def convert_md_to_docx(
    md_path: str | Path,
    docx_path: str | Path | None = None,
    reference_doc: str | Path | None = None,
) -> Path:
    """Convert a Markdown file to .docx."""
    md_path = Path(md_path)
    if docx_path is None:
        docx_path = md_path.with_suffix(".docx")
    else:
        docx_path = Path(docx_path)

    docx_path.parent.mkdir(parents=True, exist_ok=True)

    lua_path = Path(tempfile.mktemp(suffix=".lua"))
    lua_path.write_text(_MD2DOC_LUA_FILTER)

    title = _extract_title_from_frontmatter(md_path)

    # Create a temp copy with the title stripped from YAML frontmatter so
    # Pandoc doesn't render it as a visible title block in the docx.
    md_content_raw = md_path.read_text(encoding="utf-8")
    tmp_md = Path(tempfile.mktemp(suffix=".md", dir=md_path.parent))
    if md_content_raw.startswith("---"):
        end = md_content_raw.find("---", 3)
        if end != -1:
            frontmatter = md_content_raw[3:end]
            stripped_fm = "\n".join(
                ln for ln in frontmatter.splitlines()
                if not ln.strip().startswith("title:")
            )
            # Keep frontmatter block only if other keys remain.
            if stripped_fm.strip():
                md_content_raw = f"---{stripped_fm}\n---{md_content_raw[end + 3:]}"
            else:
                md_content_raw = md_content_raw[end + 3:].lstrip("\n")
    tmp_md.write_text(md_content_raw, encoding="utf-8")

    try:
        extra_args = [
            "--lua-filter", str(lua_path),
            "--standalone",
            "--from", _MD_READ_FORMAT,
            "--resource-path", str(md_path.parent),
        ]
        if reference_doc:
            extra_args += ["--reference-doc", str(Path(reference_doc).resolve())]

        pypandoc.convert_file(
            str(tmp_md),
            "docx",
            outputfile=str(docx_path),
            extra_args=extra_args,
        )
    finally:
        lua_path.unlink(missing_ok=True)
        tmp_md.unlink(missing_ok=True)

    _keep_tables_together(docx_path)

    _center_image_paragraphs(docx_path)

    # Set title in docx core properties (metadata only, not rendered visibly).
    if title:
        doc = Document(str(docx_path))
        doc.core_properties.title = title
        doc.save(str(docx_path))

    # Apply column widths derived from the markdown's own separator-line
    # proportions.  The reference doc (-s) is only used for general styling
    # (fonts, heading styles, margins) — not for table column widths.
    md_content = md_path.read_text(encoding="utf-8")
    md_proportions = _extract_md_table_proportions(md_content)
    if md_proportions:
        _apply_table_col_widths(docx_path, md_proportions)

    return docx_path


# ---------------------------------------------------------------------------
# Directory mirroring (doc2md batch mode)
# ---------------------------------------------------------------------------

def convert_directory(
    input_dir: str | Path,
    output_dir: str | Path,
) -> list[Path]:
    """Recursively convert every .docx under *input_dir* to Markdown."""
    input_dir = Path(input_dir).resolve()
    output_dir = Path(output_dir).resolve()

    if not input_dir.is_dir():
        msg = f"Input directory does not exist: {input_dir}"
        raise FileNotFoundError(msg)

    converted: list[Path] = []

    for docx_file in sorted(input_dir.rglob("*.docx")):
        if docx_file.name.startswith("~"):
            continue

        relative = docx_file.relative_to(input_dir)
        md_file = (output_dir / relative).with_suffix(".md")
        md_file.parent.mkdir(parents=True, exist_ok=True)

        try:
            with _BrailleSpinner(str(relative)):
                result = convert_docx_to_md(docx_file, md_file)
            converted.append(result)
        except Exception as exc:  # noqa: BLE001
            print(f"    {exc}", file=sys.stderr)

    return converted


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> None:
    """Entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="Bidirectional converter between .docx and Markdown.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # --- doc2md --------------------------------------------------------
    p_d2m = subparsers.add_parser(
        "doc2md", prog="doc2md", help="Convert .docx to Markdown"
    )
    p_d2m.add_argument(
        "input", nargs="?", type=Path, default=None,
        help="Input .docx file (single-file mode)",
    )
    p_d2m.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output .md file (single-file) or output directory (with -d)",
    )
    p_d2m.add_argument(
        "-d", "--input-dir", type=Path, default=None,
        help="Input directory for batch conversion (default: docx)",
    )

    # --- md2doc --------------------------------------------------------
    p_m2d = subparsers.add_parser(
        "md2doc", prog="md2doc", help="Convert Markdown to .docx"
    )
    p_m2d.add_argument(
        "input", type=Path,
        help="Input .md file",
    )
    p_m2d.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output .docx file",
    )
    p_m2d.add_argument(
        "-s", "--style-reference", type=Path, default=None,
        help="Reference .docx for styling (fonts, heading styles, margins)",
    )

    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "doc2md":
        if args.input_dir is not None or args.input is None:
            # Directory mode
            input_dir = args.input_dir or Path("docx")
            output_dir = args.output or Path("md")
            print(f"Mirroring {input_dir}/ → {output_dir}/")
            converted = convert_directory(input_dir, output_dir)
            print(f"\nDone — converted {len(converted)} file(s).")
        else:
            # Single-file mode
            if not args.input.exists():
                print(f"Error: {args.input} does not exist.", file=sys.stderr)
                sys.exit(1)
            output = args.output or args.input.with_suffix(".md")
            with _BrailleSpinner(f"{args.input} → {output}"):
                result = convert_docx_to_md(args.input, output)
            print(f"Converted: {args.input} → {result}")

    elif args.command == "md2doc":
        if not args.input.exists():
            print(f"Error: {args.input} does not exist.", file=sys.stderr)
            sys.exit(1)
        output = args.output or args.input.with_suffix(".docx")
        with _BrailleSpinner(f"{args.input} → {output}"):
            result = convert_md_to_docx(args.input, output, args.style_reference)
        print(f"Converted: {args.input} → {result}")


def doc2md() -> None:
    """Console-script entry point pinned to the ``doc2md`` subcommand."""
    main(["doc2md", *sys.argv[1:]])


def md2doc() -> None:
    """Console-script entry point pinned to the ``md2doc`` subcommand."""
    main(["md2doc", *sys.argv[1:]])


if __name__ == "__main__":
    main()
