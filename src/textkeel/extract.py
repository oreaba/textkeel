"""The front door: route a file to its reader, gate it, and write its outputs.

Office, HTML, Markdown, CSV and text files go to textkeel's own readers. PDFs and
images go through pdf-ar. A .doc is converted to .docx by a local Word on Windows,
and otherwise asks for conversion first.
"""

from __future__ import annotations

import json
import time
from pathlib import Path

from textkeel import check
from textkeel.model import Document
from textkeel.readers import (
    doc_word,
    docx_native,
    html_native,
    pdf_bridge,
    pptx_native,
    text_native,
    xlsx_native,
)

NATIVE = {".docx": docx_native.read, ".xlsx": xlsx_native.read, ".pptx": pptx_native.read,
          ".html": html_native.read, ".htm": html_native.read, ".md": text_native.read_markdown,
          ".markdown": text_native.read_markdown, ".txt": text_native.read_text,
          ".csv": text_native.read_csv}
OFFICE = (".docx", ".xlsx", ".pptx")
SUPPORTED = tuple(NATIVE) + (".pdf", ".doc") + pdf_bridge.IMAGE_SUFFIXES


def extract(path: Path, engine: str = "fonts", fonts_dir: Path = Path("fonts"),
            font_check: bool = False, ocr=None, doc_cache: Path | None = None,
            use_word: bool = True, fonts=None) -> Document:
    """``fonts``: a fonts folder already checked for this PDF (``check_pdf_fonts``)."""
    suffix = path.suffix.lower()
    started = time.perf_counter()
    if suffix == ".pdf":
        document = pdf_bridge.read_pdf(path, engine, fonts_dir, font_check, ocr, fonts)
    elif suffix in pdf_bridge.IMAGE_SUFFIXES:
        document = pdf_bridge.read_image(path, ocr)
    elif suffix == ".doc":
        document = doc_word.read(path, doc_cache, use_word)
    elif suffix in NATIVE:
        document = NATIVE[suffix](path)
    else:
        return Document(str(path), suffix.lstrip("."), "none", status="review",
                        reasons=["unsupported_format"])
    document.meta["seconds"] = round(time.perf_counter() - started, 3)
    if suffix in OFFICE:
        check.apply(document, path)
    return document


def check_pdf_fonts(files: list[Path], engine: str, fonts_dir: Path, font_check: bool,
                    root: Path | None = None, report_dir: Path | None = None):
    """Every PDF's fonts checked before any file is read: (fonts library, rows), or
    (None, []) when the engine needs no fonts. Raises ``MissingFontsError`` when a needed
    font is not there, except for ``auto``, which reads this machine's installed versions."""
    from pdf_ar.extract import check_fonts

    options = pdf_bridge.pdf_options(engine, fonts_dir, font_check)
    pdfs = [p for p in files if p.suffix.lower() == ".pdf"]
    if not pdfs or not options.uses_fonts:
        return None, []
    return check_fonts(pdfs, options, root, report_dir, strict=engine != "auto")


def write(document: Document, folder: Path) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    stem = Path(document.source).stem
    (folder / f"{stem}.json").write_text(json.dumps(document.to_dict(), ensure_ascii=False,
                                                    indent=1), encoding="utf-8")
    (folder / f"{stem}.md").write_text(document.markdown(), encoding="utf-8")
    (folder / f"{stem}.txt").write_text(document.text(("body",)), encoding="utf-8")
    with (folder / f"{stem}.chunks.jsonl").open("w", encoding="utf-8") as handle:
        for chunk in document.chunks():
            handle.write(json.dumps(chunk, ensure_ascii=False) + "\n")
