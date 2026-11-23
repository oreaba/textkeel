"""HTML reader on Python's own ``html.parser``: no extra dependency.

Kept: ``<title>``, ``h1``-``h6``, paragraphs and other text blocks, nested ``ul``/``ol``
lists, tables (``rowspan`` and ``colspan`` become spans, ``th`` marks header cells),
``pre``/``code`` blocks, ``img`` alt text and ``figcaption``. ``header``, ``footer`` and
``nav`` go to the furniture layer; ``script``, ``style`` and ``template`` are dropped.
"""

from __future__ import annotations

import re
from html.parser import HTMLParser
from pathlib import Path

from textkeel.model import Cell, Document, Item, Provenance

BLOCKS = {"p", "div", "section", "article", "main", "aside", "blockquote", "address",
          "dd", "dt", "figure", "body"}
SKIP = {"script", "style", "template", "noscript", "head"}
FURNITURE = {"header", "footer", "nav"}


class _Parser(HTMLParser):
    def __init__(self, document: Document) -> None:
        super().__init__(convert_charrefs=True)
        self.doc = document
        self.text: list[str] = []
        self.kind = "paragraph"
        self.level = 0
        self.ordered = False
        self.lists: list[bool] = []
        self.skip = 0
        self.furniture = 0
        self.pre = 0
        self.tables: list[dict] = []
        self.in_title = False
        self.title = ""

    # Text blocks -------------------------------------------------------------

    def flush(self) -> None:
        text = "".join(self.text)
        self.text = []
        text = text.strip("\n") if self.pre else re.sub(r"\s+", " ", text).strip()
        if not text:
            return
        if self.tables:
            self.tables[-1]["cell"].append(text)
            return
        layer = "furniture" if self.furniture else "body"
        kind = self.kind if not self.furniture else "paragraph"
        self.doc.add("code" if self.pre else kind, text, level=self.level, layer=layer,
                     ordered=self.ordered, prov=Provenance(part="html"))

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in SKIP:
            self.skip += 1
            return
        if tag == "title":
            self.in_title = True
            return
        if self.skip:
            return
        if tag in FURNITURE:
            self.flush()
            self.furniture += 1
        elif re.fullmatch(r"h[1-6]", tag):
            self.flush()
            self.kind, self.level = "heading", int(tag[1])
        elif tag in ("ul", "ol"):
            self.flush()
            self.lists.append(tag == "ol")
        elif tag == "li":
            self.flush()
            self.kind, self.level = "list_item", max(0, len(self.lists) - 1)
            self.ordered = bool(self.lists and self.lists[-1])
        elif tag in BLOCKS or tag == "br" and not self.tables:
            self.flush()
            if tag != "br":
                self.kind, self.level = "paragraph", 0
        elif tag == "pre":
            self.flush()
            self.pre += 1
        elif tag == "figcaption":
            self.flush()
            self.kind = "caption"
        elif tag == "img":
            alt = (attrs.get("alt") or "").strip()
            self.flush()
            self.doc.items.append(Item("picture", alt, prov=Provenance(part="html"),
                                       layer="furniture" if self.furniture else "body"))
        elif tag == "table":
            self.flush()
            self.tables.append({"cells": [], "row": -1, "occupied": set(), "cell": [],
                                "open": None})
        elif tag == "tr" and self.tables:
            table = self.tables[-1]
            table["row"] += 1
        elif tag in ("td", "th") and self.tables:
            table = self.tables[-1]
            row = max(table["row"], 0)
            col = 0
            while (row, col) in table["occupied"]:
                col += 1
            rows = int(attrs.get("rowspan") or 1)
            cols = int(attrs.get("colspan") or 1)
            for r in range(row, row + rows):
                for c in range(col, col + cols):
                    table["occupied"].add((r, c))
            table["open"] = Cell(row, col, "", rows, cols, header=tag == "th")
            table["cell"] = []

    def handle_endtag(self, tag):
        if tag in SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if tag == "title":
            self.in_title = False
            return
        if self.skip:
            return
        if tag in ("td", "th") and self.tables:
            self.flush()
            table = self.tables[-1]
            if table["open"] is not None:
                table["open"].text = "\n".join(table["cell"])
                table["cells"].append(table["open"])
                table["open"] = None
        elif tag == "table" and self.tables:
            self.flush()
            table = self.tables.pop()
            cells = table["cells"]
            rows = max((c.row + c.row_span for c in cells), default=0)
            cols = max((c.col + c.col_span for c in cells), default=0)
            item = Item("table", prov=Provenance(part="html"), cells=cells, rows=rows,
                        cols=cols, layer="furniture" if self.furniture else "body")
            if self.tables:  # a nested table becomes text in its outer cell
                self.tables[-1]["cell"].append("\n".join(c.text for c in cells if c.text))
            else:
                self.doc.items.append(item)
        elif tag in FURNITURE:
            self.flush()
            self.furniture = max(0, self.furniture - 1)
        elif re.fullmatch(r"h[1-6]", tag) or tag in BLOCKS or tag in ("li", "figcaption"):
            self.flush()
            self.kind, self.level, self.ordered = "paragraph", 0, False
            if tag == "li" and self.lists:
                self.kind, self.level = "list_item", len(self.lists) - 1
        elif tag in ("ul", "ol"):
            self.flush()
            if self.lists:
                self.lists.pop()
            self.kind, self.level = "paragraph", 0
        elif tag == "pre":
            self.flush()
            self.pre = max(0, self.pre - 1)

    def handle_data(self, data):
        if self.in_title:
            self.title += data
        elif not self.skip:
            self.text.append(data)


def read(path: Path) -> Document:
    raw = path.read_bytes()
    match = re.search(rb'charset=["\']?([\w-]+)', raw[:2048])
    encoding = match.group(1).decode() if match else "utf-8"
    try:
        html = raw.decode(encoding, errors="replace")
    except LookupError:
        html = raw.decode("utf-8", errors="replace")
    document = Document(str(path), "html", "native")
    parser = _Parser(document)
    parser.feed(html)
    parser.flush()
    if parser.title.strip():
        document.items.insert(0, Item("title", parser.title.strip(), prov=Provenance(part="html")))
    return document
