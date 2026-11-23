"""Word (.docx) reader that works straight from the file's XML.

Reading order is the body's own order. Kept:

- titles and headings (style Title, Heading N, or an outline level);
- numbered and bulleted list items with their level;
- tables, with ``gridSpan`` and ``vMerge`` turned into row and column spans and each
  merged cell stored once;
- text boxes (drawn shapes), each paragraph after its anchor paragraph;
- pictures, with their alt text;
- page headers and footers (every distinct variant once) as the furniture layer;
- footnotes and endnotes.

Text in ``mc:Fallback`` (a duplicate kept for old Word versions) is skipped.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from textkeel.model import Cell, Document, Item, Provenance

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
HEADING = re.compile(r"^heading\s*(\d)$", re.I)


def _styles(z: zipfile.ZipFile) -> dict[str, tuple[str, int | None, tuple | None]]:
    """styleId -> (name, outline level, (numId, level) list numbering), following basedOn."""
    if "word/styles.xml" not in z.namelist():
        return {}
    root = ET.fromstring(z.read("word/styles.xml"))
    raw = {}
    for style in root.iter(W + "style"):
        sid = style.get(W + "styleId")
        name = style.find(W + "name")
        level = style.find(f"{W}pPr/{W}outlineLvl")
        based = style.find(W + "basedOn")
        num = style.find(f"{W}pPr/{W}numPr/{W}numId")
        ilvl = style.find(f"{W}pPr/{W}numPr/{W}ilvl")
        raw[sid] = (name.get(W + "val") if name is not None else sid,
                    int(level.get(W + "val")) if level is not None else None,
                    based.get(W + "val") if based is not None else None,
                    (num.get(W + "val"), int(ilvl.get(W + "val")) if ilvl is not None else 0)
                    if num is not None else None)
    out = {}
    for sid, (name, level, based, numbering) in raw.items():
        seen = {sid}
        while (level is None or numbering is None) and based in raw and based not in seen:
            seen.add(based)
            level = level if level is not None else raw[based][1]
            numbering = numbering or raw[based][3]
            based = raw[based][2]
        out[sid] = (name or sid, level, numbering)
    return out


def _numbering(z: zipfile.ZipFile) -> dict[tuple[str, int], bool]:
    """(numId, level) -> numbered (True) or bulleted (False)."""
    if "word/numbering.xml" not in z.namelist():
        return {}
    root = ET.fromstring(z.read("word/numbering.xml"))
    abstract = {}
    for a in root.iter(W + "abstractNum"):
        levels = {}
        for lvl in a.iter(W + "lvl"):
            fmt = lvl.find(W + "numFmt")
            levels[int(lvl.get(W + "ilvl"))] = (fmt is not None
                                                and fmt.get(W + "val") not in ("bullet", "none"))
        abstract[a.get(W + "abstractNumId")] = levels
    out = {}
    for num in root.iter(W + "num"):
        ref = num.find(W + "abstractNumId")
        if ref is None:
            continue
        for level, ordered in abstract.get(ref.get(W + "val"), {}).items():
            out[(num.get(W + "numId"), level)] = ordered
    return out


def _runs_text(node: ET.Element) -> str:
    """A paragraph's own text: runs, tabs and breaks, not text boxes or fallbacks."""
    out: list[str] = []

    def walk(el: ET.Element) -> None:
        if el.tag in (W + "txbxContent", MC_FALLBACK):
            return
        if el.tag == W + "t":
            out.append(el.text or "")
        elif el.tag in (W + "tab", W + "br", W + "cr"):
            out.append(" ")
        for child in el:
            walk(child)

    walk(node)
    return re.sub(r"[ \t]+", " ", "".join(out)).strip()


def _paragraphs(node: ET.Element):
    """Every paragraph under ``node``, text-box paragraphs included, fallbacks skipped."""
    for child in node:
        if child.tag == MC_FALLBACK:
            continue
        if child.tag == W + "p":
            yield child
        yield from _paragraphs(child)


class _Reader:
    def __init__(self, z: zipfile.ZipFile, document: Document) -> None:
        self.styles = _styles(z)
        self.numbering = _numbering(z)
        self.doc = document

    def paragraph(self, p: ET.Element, layer: str, part: str) -> None:
        text = _runs_text(p)
        ppr = p.find(W + "pPr")
        kind, level, ordered = "paragraph", 0, False
        if ppr is not None:
            style = ppr.find(W + "pStyle")
            name, outline, style_numbering = self.styles.get(
                style.get(W + "val"), ("", None, None)) if style is not None else ("", None, None)
            own = ppr.find(W + "outlineLvl")
            if own is not None:
                outline = int(own.get(W + "val"))
            heading = HEADING.match(name or "")
            if (name or "").lower() == "title":
                kind = "title"
            elif heading or (outline is not None and outline < 9):
                kind, level = "heading", int(heading.group(1)) if heading else outline + 1
            numpr = ppr.find(W + "numPr")
            num_id, num_level = style_numbering or (None, 0)
            if numpr is not None:
                ilvl, num = numpr.find(W + "ilvl"), numpr.find(W + "numId")
                num_id = num.get(W + "val") if num is not None else num_id
                num_level = int(ilvl.get(W + "val")) if ilvl is not None else num_level
            if kind == "paragraph" and num_id not in (None, "0"):
                kind, level = "list_item", num_level
                ordered = self.numbering.get((num_id, level), False)
            if (name or "").lower() in ("caption",):
                kind = "caption"
        if kind != "paragraph" and layer != "body":
            kind, level = "paragraph", 0
        prov = Provenance(part=part)
        self.doc.add(kind, text, level=level, layer=layer, prov=prov, ordered=ordered)
        for drawing in p.iter(W + "drawing"):
            if drawing.find(f".//{A_BLIP}") is None:
                continue  # a drawn shape or text box, not a picture
            props = drawing.find(f".//{WP}docPr")
            alt = ((props.get("descr") or props.get("title") or "") if props is not None
                   else "").strip()
            self.doc.items.append(Item("picture", alt, layer=layer, prov=Provenance(part=part)))
        for box in p.iter(W + "txbxContent"):
            if not self._in_fallback(p, box):
                self.block(box, layer, "textbox")

    @staticmethod
    def _in_fallback(top: ET.Element, target: ET.Element) -> bool:
        stack = [(top, False)]
        while stack:
            node, inside = stack.pop()
            if node is target:
                return inside
            for child in node:
                stack.append((child, inside or child.tag == MC_FALLBACK))
        return False

    def table(self, tbl: ET.Element, layer: str, part: str) -> None:
        cells: list[Cell] = []
        open_cells: dict[int, Cell] = {}  # column -> cell still spanning down (vMerge)
        rows = 0
        cols = 0
        for r, tr in enumerate(tbl.findall(W + "tr")):
            rows = r + 1
            col = 0
            header = tr.find(f"{W}trPr/{W}tblHeader") is not None
            for tc in tr.findall(W + "tc"):
                pr = tc.find(W + "tcPr")
                span_el = pr.find(W + "gridSpan") if pr is not None else None
                span = int(span_el.get(W + "val")) if span_el is not None else 1
                vmerge = pr.find(W + "vMerge") if pr is not None else None
                text = "\n".join(t for t in (_runs_text(p) for p in _paragraphs(tc)) if t)
                if vmerge is not None and vmerge.get(W + "val") != "restart" and col in open_cells:
                    open_cells[col].row_span += 1
                else:
                    cell = Cell(r, col, text, 1, span, header)
                    cells.append(cell)
                    if vmerge is not None:
                        open_cells[col] = cell
                    else:
                        open_cells.pop(col, None)
                col += span
            cols = max(cols, col)
        self.doc.items.append(Item("table", layer=layer, prov=Provenance(part=part),
                                   cells=cells, rows=rows, cols=cols))

    def block(self, container: ET.Element, layer: str, part: str) -> None:
        for child in container:
            if child.tag == W + "p":
                self.paragraph(child, layer, part)
            elif child.tag == W + "tbl":
                self.table(child, layer, part)
            elif child.tag == W + "sdt":
                content = child.find(W + "sdtContent")
                if content is not None:
                    self.block(content, layer, part)


def _relationships(z: zipfile.ZipFile) -> dict[str, str]:
    name = "word/_rels/document.xml.rels"
    if name not in z.namelist():
        return {}
    return {rel.get("Id"): "word/" + rel.get("Target").lstrip("/").removeprefix("word/")
            for rel in ET.fromstring(z.read(name)).iter(REL)}


def read(path: Path) -> Document:
    document = Document(str(path), "docx", "native")
    with zipfile.ZipFile(path) as z:
        reader = _Reader(z, document)
        root = ET.fromstring(z.read("word/document.xml"))
        body = root.find(W + "body")
        rels = _relationships(z)
        # Headers and footers referenced by the sections, each distinct text once.
        furniture: list[tuple[str, str]] = []
        for ref in root.iter():
            if ref.tag in (W + "headerReference", W + "footerReference"):
                target = rels.get(ref.get(R + "id"))
                kind = "header" if ref.tag == W + "headerReference" else "footer"
                if target and (kind, target) not in furniture and target in z.namelist():
                    furniture.append((kind, target))
        seen_text: set[str] = set()
        start = len(document.items)
        for kind, target in furniture:
            if kind != "header":
                continue
            before = len(document.items)
            reader.block(ET.fromstring(z.read(target)), "furniture", "header")
            _drop_repeats(document, before, seen_text)
        header_items = document.items[start:]
        del document.items[start:]
        if body is not None:
            reader.block(body, "body", "body")
        document.items[start:start] = header_items
        for kind, target in furniture:
            if kind == "footer":
                before = len(document.items)
                reader.block(ET.fromstring(z.read(target)), "furniture", "footer")
                _drop_repeats(document, before, seen_text)
        for name in ("word/footnotes.xml", "word/endnotes.xml"):
            if name in z.namelist():
                for note in ET.fromstring(z.read(name)):
                    if note.get(W + "type") in ("separator", "continuationSeparator"):
                        continue
                    for p in note.iter(W + "p"):
                        document.add("footnote", _runs_text(p), prov=Provenance(part="footnote"))
    return document


def _drop_repeats(document: Document, start: int, seen: set[str]) -> None:
    """Keep a header or footer variant only if its text differs from one already kept."""
    added = document.items[start:]
    key = "\n".join(i.text or "\n".join(c.text for c in i.cells) for i in added)
    if key in seen:
        del document.items[start:]
    else:
        seen.add(key)
