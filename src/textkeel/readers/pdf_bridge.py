"""PDF and image readers: pdf-ar underneath, its blocks turned into the shared model.

PDFs go through pdf-ar's own page chain (``fonts``, ``native``, ``ocr`` or ``auto``),
unchanged; drawn table cells are grouped into tables with row and column spans (the
same grouping pdf-ar's Markdown uses), drawn boxes become paragraphs from part
``shape``. Page status and reasons are kept per page in ``meta["pages"]``.

Images (.png, .jpg, .tif, ...) are wrapped as a one-page PDF by PyMuPDF (WebP through
Pillow first) and read by pdf-ar's OCR engine.
"""

from __future__ import annotations

import re
import tempfile
from collections import Counter
from pathlib import Path

from textkeel.model import Cell, Document, Item, Provenance

IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp")


def _table(blocks, page: int) -> Item:
    from pdf_ar.formats import EDGE_TOLERANCE, _boundaries, _index

    columns = _boundaries([b.bbox[0] for b in blocks] + [b.bbox[2] for b in blocks])
    rows = _boundaries([b.bbox[1] for b in blocks] + [b.bbox[3] for b in blocks])
    rtl = sum(b.rtl for b in blocks) * 2 >= len(blocks)
    n_rows, n_cols = max(1, len(rows) - 1), max(1, len(columns) - 1)
    cells = []
    for block in blocks:
        x0, y0, x1, y1 = block.bbox
        r0, r1 = _index(rows, y0), _index(rows, y1)
        c0, c1 = _index(columns, x0), _index(columns, x1)
        row_span, col_span = max(1, r1 - r0), max(1, c1 - c0)
        col = n_cols - c0 - col_span if rtl else c0
        cells.append(Cell(r0, max(0, col), block.text, row_span, col_span))
    x0 = min(b.bbox[0] for b in blocks) - EDGE_TOLERANCE
    return Item("table", prov=Provenance(page=page, bbox=(
        x0, min(b.bbox[1] for b in blocks), max(b.bbox[2] for b in blocks),
        max(b.bbox[3] for b in blocks)), part="pdf"), cells=cells, rows=n_rows, cols=n_cols)


LIST_MARKER = re.compile(r"^([▪•·◦●○■□➢✓\-–o]|\(?\d{1,2}[.)]|\(?[\u0621-\u064aA-Za-z][.)])\s+")
NUMBERED_SECTION = re.compile(r"^(\d{1,2}(?:\.\d{1,2})*)\.?\s+\S")
BOLD_NAME = re.compile(r"bold|black|heavy|semibold|demi", re.I)
TOC_LEADER = re.compile(r"\.{5,}|…{2,}")


def _span_styles(page) -> list[tuple[tuple, float, bool, int]]:
    """(bbox, size, bold, characters) of every visible text span on the page."""
    out = []
    for span in page.spans():
        if span.get("type") == 3 or span.get("opacity", 1) == 0:
            continue
        chars = sum(1 for c in span["chars"] if c[0] > 32)
        if chars:
            bold = bool(span.get("flags", 0) & 16) or bool(BOLD_NAME.search(span["font"]))
            out.append((span["bbox"], round(span["size"] * 2) / 2, bold, chars))
    return out


def _block_style(block, styles) -> tuple[float, float]:
    """The block's dominant text size and the share of its characters set in bold."""
    x0, y0, x1, y1 = block.bbox
    inside = [s for s in styles if x0 - 2 <= (s[0][0] + s[0][2]) / 2 <= x1 + 2
              and y0 - 2 <= (s[0][1] + s[0][3]) / 2 <= y1 + 2]
    if not inside:
        return 0.0, 0.0
    sizes = Counter()
    for _box, size, _bold, chars in inside:
        sizes[size] += chars
    total = sum(chars for *_r, chars in inside)
    return sizes.most_common(1)[0][0], sum(c for _b, _s, bold, c in inside if bold) / total


def _classify(entries: list[dict]) -> None:
    """Headings and list items for PDF paragraphs, from text size, weight and markers.

    The body size is the size most characters of the document are set in. A short
    paragraph (at most 14 words, not ending a sentence) is a heading when its text is at
    least 15% larger, or when it is all bold; larger sizes get higher levels, and a
    numbered section such as "2.1" takes its depth from its numbering. Bulleted lines and
    table-of-contents lines (dot leaders) are never headings. A drawn box can be a heading
    only as a wide section bar (at least 250 pt) of at most 8 words. Paragraphs that start with
    a bullet or a list number are list items, nested by their indent from the page's
    leftmost list item.
    """
    weights = Counter()
    for entry in entries:
        weights[entry["size"]] += entry["chars"]
    if not weights:
        return
    body = weights.most_common(1)[0][0]
    candidates = []
    for entry in entries:
        text = entry["item"].text
        words = len(text.split())
        short = 0 < words <= 14 and not text.rstrip().endswith((".", "؛", ",", "،"))
        large = entry["size"] >= body * 1.15
        marker = LIST_MARKER.match(text)
        bullet = marker is not None and not NUMBERED_SECTION.match(text)
        contents_line = TOC_LEADER.search(text) is not None
        # A shaded section bar is a drawn box too: wide, with a few bold words. Narrow
        # boxes (flowchart steps) never become headings.
        bar = entry["part"] == "shape" and entry["width"] >= 250 and words <= 8
        if ((entry["part"] == "pdf" or bar) and short and (large or entry["bold"] >= 0.9)
                and not bullet and not contents_line):
            candidates.append(entry)
    heading_sizes = sorted({e["size"] for e in candidates if e["size"] >= body * 1.15},
                           reverse=True)
    for entry in candidates:
        item = entry["item"]
        numbered = NUMBERED_SECTION.match(item.text)
        if numbered:
            level = numbered.group(1).count(".") + 1
        elif entry["size"] in heading_sizes:
            level = heading_sizes.index(entry["size"]) + 1
        else:
            level = len(heading_sizes) + 1
        item.kind, item.level = "heading", min(level, 6)
    items = [e for e in entries if e["item"].kind == "paragraph" and e["part"] == "pdf"
             and LIST_MARKER.match(e["item"].text)]
    margins: dict[int, float] = {}
    for entry in items:
        page = entry["item"].prov.page
        margins[page] = min(margins.get(page, entry["x0"]), entry["x0"])
    for entry in items:
        item = entry["item"]
        marker = LIST_MARKER.match(item.text)
        item.kind = "list_item"
        item.ordered = marker.group(1)[0] not in "▪•·◦●○■□➢✓-–o"
        item.text = item.text[marker.end():].strip() or item.text
        indent = abs(entry["x0"] - margins.get(item.prov.page, entry["x0"]))
        item.level = min(3, int(indent // 18))


def pdf_options(engine: str = "fonts", fonts_dir: Path = Path("fonts"), font_check: bool = False,
                ocr=None):
    from pdf_ar.extract import Options
    from pdf_ar.ocr import OcrOptions

    return Options(out_dir=Path(tempfile.gettempdir()), fonts_dir=fonts_dir, engine=engine,
                   font_check=font_check, ocr=ocr or OcrOptions())


def read_pdf(path: Path, engine: str = "fonts", fonts_dir: Path = Path("fonts"),
             font_check: bool = False, ocr=None, library=None) -> Document:
    """``library``: a fonts folder already checked for this PDF; otherwise the PDF's fonts
    are checked first and ``MissingFontsError`` stops before anything is read, except for
    ``auto``, which reads this machine's installed versions of missing fonts."""
    from pdf_ar.extract import _page, check_fonts
    from pdf_ar.formats import _tables
    from pdf_ar.glyphs import FontMapper
    from pdf_ar.pdf_adapter import PdfDocument

    options = pdf_options(engine, fonts_dir, font_check, ocr)
    if options.uses_fonts and library is None:
        library, _rows = check_fonts([path], options, strict=engine != "auto")
    elif not options.uses_fonts:
        library = None
    document = Document(str(path), "pdf", f"pdf-ar:{engine}")
    pages, entries = [], []
    with PdfDocument(path) as pdf:
        mapper = FontMapper(pdf, library)
        document.pages = pdf.page_count
        for number in range(1, pdf.page_count + 1):
            page = pdf.page(number)
            (report, _text, _lines, blocks), _ocr = _page(page, number, mapper, options,
                                                          page.has_images())
            pages.append({"page": number, "status": report.status, "reasons": report.reasons,
                          "source": report.source})
            styles = _span_styles(page) if report.source != "ocr" else []
            groups = {g[0]: g for g in _tables(blocks) if len(g) > 1}
            in_table = {i for g in groups.values() for i in g}
            for index, block in enumerate(blocks):
                if index in groups:
                    document.items.append(_table([blocks[i] for i in groups[index]], number))
                elif index not in in_table:
                    part = "shape" if block.kind == "box" else ("cell" if block.kind == "cell"
                                                                 else "pdf")
                    item = document.add("paragraph", block.text,
                                        prov=Provenance(page=number, bbox=block.bbox, part=part))
                    if item is not None and styles:
                        size, bold = _block_style(block, styles)
                        entries.append({"item": item, "size": size, "bold": bold, "part": part,
                                        "chars": len(block.text), "x0": block.bbox[0]
                                        if not block.rtl else -block.bbox[2],
                                        "width": block.bbox[2] - block.bbox[0]})
    _classify(entries)
    document.meta["pages"] = pages
    reasons = Counter(r for p in pages for r in p["reasons"])
    document.status = "ok" if all(p["status"] == "ok" for p in pages) else "review"
    document.reasons = sorted(reasons)
    return document


def read_image(path: Path, ocr=None) -> Document:
    import pymupdf
    from pdf_ar.ocr import OcrOptions

    with tempfile.TemporaryDirectory() as tmp:
        wrapped = Path(tmp) / "image.pdf"
        try:
            image = pymupdf.open(path)
        except pymupdf.FileDataError:
            # PyMuPDF cannot open WebP; Pillow comes with pytesseract, which OCR needs anyway.
            from PIL import Image

            png = Path(tmp) / "image.png"
            with Image.open(path) as source:
                source.save(png, dpi=source.info.get("dpi", (300, 300)))
            image = pymupdf.open(png)
        wrapped.write_bytes(image.convert_to_pdf())
        image.close()
        document = read_pdf(wrapped, engine="ocr", ocr=ocr or OcrOptions())
    document.source, document.format = str(path), path.suffix.lower().lstrip(".")
    return document
