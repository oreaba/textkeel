"""Markdown, plain text and CSV readers, on the standard library only."""

from __future__ import annotations

import csv
import io
import re
from pathlib import Path

from textkeel.model import Cell, Document, Item, Provenance

LIST = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*$")
TABLE_RULE = re.compile(r"^\s*\|?\s*:?-{3,}")


def _decode(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1256", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _row(line: str) -> list[str]:
    line = line.strip().strip("|")
    return [c.strip() for c in re.split(r"(?<!\\)\|", line)]


def read_markdown(path: Path) -> Document:
    document = Document(str(path), "md", "native")
    lines = _decode(path).splitlines()
    paragraph: list[str] = []
    prov = Provenance(part="md")

    def flush() -> None:
        if paragraph:
            document.add("paragraph", " ".join(paragraph), prov=prov)
            paragraph.clear()

    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("```"):
            flush()
            block = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                block.append(lines[i])
                i += 1
            document.add("code", "\n".join(block), prov=prov)
        elif HEADING.match(line):
            flush()
            hashes, text = HEADING.match(line).groups()
            document.add("heading", text, level=len(hashes), prov=prov)
        elif "|" in line and i + 1 < len(lines) and TABLE_RULE.match(lines[i + 1]):
            flush()
            rows = [_row(line)]
            i += 2
            while i < len(lines) and "|" in lines[i] and lines[i].strip():
                rows.append(_row(lines[i]))
                i += 1
            cols = max(len(r) for r in rows)
            cells = [Cell(r, c, text, header=r == 0) for r, row in enumerate(rows)
                     for c, text in enumerate(row) if text]
            document.items.append(Item("table", prov=prov, cells=cells, rows=len(rows),
                                       cols=cols))
            continue
        elif LIST.match(line):
            flush()
            indent, marker, text = LIST.match(line).groups()
            document.add("list_item", text, level=len(indent.expandtabs(4)) // 2,
                         ordered=marker[0].isdigit(), prov=prov)
        elif not line.strip():
            flush()
        else:
            paragraph.append(line.strip())
        i += 1
    flush()
    return document


def read_text(path: Path) -> Document:
    document = Document(str(path), "txt", "native")
    for block in re.split(r"\n\s*\n", _decode(path)):
        document.add("paragraph", re.sub(r"\s*\n\s*", " ", block), prov=Provenance(part="txt"))
    return document


def read_csv(path: Path) -> Document:
    text = _decode(path)
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(io.StringIO(text), dialect))
    document = Document(str(path), "csv", "native")
    cells = [Cell(r, c, value.strip(), header=r == 0) for r, row in enumerate(rows)
             for c, value in enumerate(row) if value.strip()]
    document.items.append(Item("table", prov=Provenance(part="csv"), cells=cells,
                               rows=len(rows), cols=max((len(r) for r in rows), default=0)))
    return document
