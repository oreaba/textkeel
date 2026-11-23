"""Office quality gate: the extracted words against the text the file itself stores.

A .docx, .xlsx or .pptx is a zip of XML parts holding its text exactly, so every
extraction can be checked deterministically: per part (body, table, text box, header,
footer, footnote, cell, slide), the share of stored words the extraction kept, and
overall, the share of extracted words that are stored ones (repeats lower it).

The reference here is deliberately independent of the readers: it walks every text
run (``w:t``, ``a:t``) or cell value, skipping ``mc:Fallback`` duplicates.
"""

from __future__ import annotations

import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
STRIP = "\"'“”‘’()[]{}<>«»,.;:!?؟،؛*-–—_/\\|#"
MIN_RECALL, MIN_PRECISION = 0.98, 0.98


def words(text: str) -> Counter:
    out: Counter = Counter()
    for token in text.split():
        token = unicodedata.normalize("NFKC", token)
        token = "".join(c for c in token if not unicodedata.category(c).startswith("M"))
        token = token.replace("\u0640", "").strip(STRIP)
        if token:
            out[token] += 1
    return out


def _texts(root: ET.Element, part: str, tag: str) -> list[tuple[str, str]]:
    """(part, paragraph text): the runs of each paragraph joined, as Word shows them."""
    paragraph = W + "p" if tag == W + "t" else A + "p"
    out = []

    def own_text(p: ET.Element) -> str:
        pieces = []

        def collect(node):
            if node.tag in (MC_FALLBACK, W + "txbxContent") or (node is not p
                                                                 and node.tag == paragraph):
                return
            if node.tag == tag and node.text:
                pieces.append(node.text)
            elif node.tag in (W + "tab", W + "br", A + "br"):
                pieces.append(" ")
            for child in node:
                collect(child)

        collect(p)
        return "".join(pieces)

    def walk(node, current):
        if node.tag == MC_FALLBACK:
            return
        if node.tag == W + "txbxContent":
            current = "textbox"
        elif node.tag == W + "tbl" and current == "body":
            current = "table"
        if node.tag == paragraph:
            text = own_text(node)
            if text.strip():
                out.append((current, text))
        for child in node:
            walk(child, current)

    walk(root, part)
    return out


def stored_text(path: Path) -> list[tuple[str, str]]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        with zipfile.ZipFile(path) as z:
            out = _texts(ET.fromstring(z.read("word/document.xml")), "body", W + "t")
            seen: set[str] = set()
            for name in sorted(z.namelist()):
                for prefix, part in (("word/header", "header"), ("word/footer", "footer"),
                                     ("word/footnotes", "footnote"), ("word/endnotes", "footnote")):
                    if name.startswith(prefix) and name.endswith(".xml"):
                        texts = _texts(ET.fromstring(z.read(name)), part, W + "t")
                        key = "".join(t for _p, t in texts)
                        if part in ("header", "footer") and key in seen:
                            continue  # an identical variant (first page, even pages)
                        seen.add(key)
                        out += [(part, t) for _p, t in texts]
            return out
    if suffix == ".xlsx":
        import openpyxl

        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
        out = []
        for sheet in book.worksheets:
            out.append(("sheet", sheet.title))
            for row in sheet.iter_rows(values_only=True):
                for value in row:
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    if value is not None and str(value).strip():
                        out.append(("cell", str(value)))
        book.close()
        return out
    if suffix == ".pptx":
        out = []
        with zipfile.ZipFile(path) as z:
            for name in sorted(z.namelist()):
                if re.match(r"ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml$", name):
                    part = "slide" if "/slides/" in name else "notes"
                    out += [(part, t) for _p, t in _texts(ET.fromstring(z.read(name)), part,
                                                          A + "t")]
        return out
    raise ValueError(f"no stored-text reference for {suffix}")


def score(document_text: str, path: Path) -> dict:
    stored = stored_text(path)
    reference = words("\n".join(t for _p, t in stored))
    found = words(document_text)
    common = sum((found & reference).values())
    recall = common / max(1, sum(reference.values()))
    precision = common / max(1, sum(found.values())) if found else 0.0
    pool = Counter(found)
    by_part: dict[str, Counter] = defaultdict(Counter)
    for part, text in stored:
        by_part[part].update(words(text))
    parts = {}
    for part in ("body", "table", "textbox", "header", "footer", "footnote", "cell", "sheet",
                 "slide", "notes"):
        if part in by_part:
            hit = by_part[part] & pool
            pool -= hit
            parts[part] = round(sum(hit.values()) / sum(by_part[part].values()), 4)
    reasons = []
    if recall < MIN_RECALL:
        reasons.append("lost_words")
    if precision < MIN_PRECISION:
        reasons.append("repeated_words")
    return {"recall": round(recall, 4), "precision": round(precision, 4),
            "f1": round(2 * recall * precision / (recall + precision), 4)
            if recall + precision else 0.0,
            "parts": parts, "reasons": reasons, "stored_words": sum(reference.values())}


def apply(document, path: Path) -> None:
    """Score ``document`` against the Office file at ``path``; failures send it to review."""
    gate = score(document.text(), path)
    document.meta["office_check"] = gate
    if gate["reasons"]:
        document.status = "review"
        document.reasons = sorted(set(document.reasons) | set(gate["reasons"]))
