"""PowerPoint (.pptx) reader: slides as pages, shapes in reading order.

Shapes are read top to bottom, then left to right (the file's own order is the order
they were drawn in, not read in). The title placeholder becomes a heading; paragraphs
with a bullet level become list items; tables keep merged cells once; groups are read
inside; pictures keep their alt text; speaker notes go to the ``notes`` layer.
"""

from __future__ import annotations

from pathlib import Path

from textkeel.model import Cell, Document, Item, Provenance

EMU_PER_POINT = 12700


def _bbox(shape):
    if shape.left is None or shape.top is None:
        return None
    x, y = shape.left / EMU_PER_POINT, shape.top / EMU_PER_POINT
    return (x, y, x + (shape.width or 0) / EMU_PER_POINT, y + (shape.height or 0) / EMU_PER_POINT)


def _shapes(document: Document, shapes, slide: int, layer: str) -> None:
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

    ordered = sorted(shapes, key=lambda s: ((s.top or 0) // 20000, s.left or 0))
    for shape in ordered:
        prov = Provenance(page=slide, bbox=_bbox(shape), part="slide" if layer == "body"
                          else layer)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            _shapes(document, shape.shapes, slide, layer)
            continue
        if getattr(shape, "has_table", False) and shape.has_table:
            table = shape.table
            cells = []
            for r, row in enumerate(table.rows):
                for c, cell in enumerate(row.cells):
                    if cell.is_spanned:
                        continue
                    cells.append(Cell(r, c, cell.text.strip(),
                                      cell.span_height if cell.is_merge_origin else 1,
                                      cell.span_width if cell.is_merge_origin else 1,
                                      header=r == 0 and table.first_row))
            document.items.append(Item("table", layer=layer, prov=prov, cells=cells,
                                       rows=len(table.rows), cols=len(table.columns)))
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            alt = shape._element.xpath("./*[1]/*[1]/@descr")
            document.items.append(Item("picture", (alt[0] if alt else "").strip(), layer=layer,
                                       prov=prov))
            continue
        if not getattr(shape, "has_text_frame", False) or not shape.has_text_frame:
            continue
        is_title = shape.is_placeholder and shape.placeholder_format.type in (
            PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
        for para in shape.text_frame.paragraphs:
            text = para.text.replace("\v", " ").strip()  # runs, fields and line breaks
            if is_title:
                document.add("heading", text, level=1, layer=layer, prov=prov)
            elif para.level > 0:
                document.add("list_item", text, level=para.level - 1, layer=layer, prov=prov)
            else:
                document.add("paragraph", text, layer=layer, prov=prov)


def read(path: Path) -> Document:
    from pptx import Presentation

    document = Document(str(path), "pptx", "native")
    deck = Presentation(str(path))
    for number, slide in enumerate(deck.slides, 1):
        _shapes(document, slide.shapes, number, "body")
        if slide.has_notes_slide:
            _shapes(document, slide.notes_slide.shapes, number, "notes")
    document.pages = len(deck.slides)
    return document
