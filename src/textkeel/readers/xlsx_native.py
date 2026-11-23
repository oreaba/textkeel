"""Excel (.xlsx) reader: each sheet a heading and a table, merged cells once.

Values are those stored in the file (formulas give their last cached result). The
table covers the sheet's non-empty area; a merged range becomes one cell with its row
and column spans. Hidden sheets are kept; their state is in the provenance part.
"""

from __future__ import annotations

from pathlib import Path

from textkeel.model import Cell, Document, Item, Provenance


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def read(path: Path) -> Document:
    import openpyxl

    document = Document(str(path), "xlsx", "native")
    book = openpyxl.load_workbook(path, data_only=True)
    for number, sheet in enumerate(book.worksheets, 1):
        state = "" if sheet.sheet_state == "visible" else f":{sheet.sheet_state}"
        document.add("heading", sheet.title, level=1,
                     prov=Provenance(page=number, part=f"sheet:{sheet.title}{state}"))
        values = {(c.row, c.column): _text(c.value) for row in sheet.iter_rows() for c in row
                  if _text(c.value)}
        if not values:
            continue
        top, left = min(r for r, _ in values), min(c for _, c in values)
        bottom, right = max(r for r, _ in values), max(c for _, c in values)
        merged = {}
        covered = set()
        for rng in sheet.merged_cells.ranges:
            merged[(rng.min_row, rng.min_col)] = (rng.max_row - rng.min_row + 1,
                                                  rng.max_col - rng.min_col + 1)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        covered.add((r, c))
        cells = []
        for (r, c), text in sorted(values.items()):
            if (r, c) in covered:
                continue
            rows, cols = merged.get((r, c), (1, 1))
            cells.append(Cell(r - top, c - left, text, rows, cols, header=r == top))
        document.items.append(Item("table", prov=Provenance(page=number,
                                                            part=f"sheet:{sheet.title}"),
                                   cells=cells, rows=bottom - top + 1, cols=right - left + 1))
    document.pages = len(book.worksheets)
    return document
