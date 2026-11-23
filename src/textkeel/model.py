"""The structured document every reader produces, and its exports.

Close in spirit to Docling's ``DoclingDocument``, but small and deterministic:

- a document is an ordered list of items in reading order;
- each item has a kind (``title``, ``heading``, ``paragraph``, ``list_item``, ``table``,
  ``picture``, ``caption``, ``footnote``, ``code``), a content layer (``body``,
  ``furniture`` for page headers and footers, ``notes`` for speaker notes), and
  provenance (page or sheet or slide, position when known, the source part);
- a table holds a grid of cells, each with its row, column, spans and header flag;
  a merged cell is stored once, at its top-left position.

Exports: JSON (``to_dict``), Markdown (merged cells written once, the rest left empty),
plain text, and heading-aware chunks for retrieval.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field

KINDS = ("title", "heading", "paragraph", "list_item", "table", "picture", "caption",
         "footnote", "code")
LAYERS = ("body", "furniture", "notes")


@dataclass
class Provenance:
    page: int | None = None  # page, slide or sheet number, 1-based
    bbox: tuple[float, float, float, float] | None = None  # points, top-left origin
    part: str | None = None  # e.g. "body", "header", "footer", "textbox", "sheet:Name"


@dataclass
class Cell:
    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1
    header: bool = False


@dataclass
class Item:
    kind: str
    text: str = ""
    level: int = 0  # heading level (1 = top) or list nesting (0 = top)
    layer: str = "body"
    prov: Provenance = field(default_factory=Provenance)
    cells: list[Cell] = field(default_factory=list)  # tables only
    rows: int = 0
    cols: int = 0
    ordered: bool = False  # numbered list item

    def __post_init__(self) -> None:
        if self.kind not in KINDS:
            raise ValueError(f"unknown kind {self.kind!r}")
        if self.layer not in LAYERS:
            raise ValueError(f"unknown layer {self.layer!r}")


@dataclass
class Document:
    source: str
    format: str
    reader: str
    items: list[Item] = field(default_factory=list)
    pages: int | None = None
    status: str = "ok"
    reasons: list[str] = field(default_factory=list)
    meta: dict = field(default_factory=dict)

    def add(self, kind: str, text: str = "", **kwargs) -> Item | None:
        if kind != "table" and kind != "picture" and not text.strip():
            return None
        item = Item(kind, text.strip(), **kwargs)
        self.items.append(item)
        return item

    # Exports -----------------------------------------------------------------

    def to_dict(self) -> dict:
        return asdict(self)

    def text(self, layers: tuple[str, ...] = LAYERS) -> str:
        """Plain text: every item's text, table cells once each, in reading order.

        A picture's alt text is metadata, not text on the page, so it is left out.
        """
        out = []
        for item in self.items:
            if item.layer not in layers or item.kind == "picture":
                continue
            if item.kind == "table":
                out.extend(c.text for c in item.cells if c.text)
            elif item.text:
                out.append(item.text)
        return "\n".join(out)

    def markdown(self, layers: tuple[str, ...] = ("body",)) -> str:
        out: list[str] = []
        page = None
        for item in self.items:
            if item.layer not in layers:
                continue
            if item.prov.page is not None and item.prov.page != page and self.pages:
                page = item.prov.page
                out.append(f"<!-- page {page} -->")
            if item.kind == "title":
                out.append(f"# {item.text}")
            elif item.kind == "heading":
                out.append(f"{'#' * min(6, max(1, item.level))} {item.text}")
            elif item.kind == "list_item":
                marker = "1." if item.ordered else "-"
                out.append(f"{'  ' * item.level}{marker} {item.text}")
            elif item.kind == "table":
                out.append(markdown_table(item))
            elif item.kind == "picture":
                out.append(f"<!-- picture: {item.text} -->" if item.text else "<!-- picture -->")
            elif item.kind == "code":
                out.append(f"```\n{item.text}\n```")
            else:
                out.append(item.text)
        return "\n\n".join(out) + "\n"

    def chunks(self, max_chars: int = 1500) -> list[dict]:
        """Body text split at headings and at ``max_chars``, each chunk with its headings."""
        chunks: list[dict] = []
        path: list[tuple[int, str]] = []  # (level, heading); a title is level 0
        buffer: list[str] = []
        pages: set[int] = set()

        def flush() -> None:
            if buffer:
                chunks.append({"headings": [text for _level, text in path],
                               "text": "\n".join(buffer), "pages": sorted(pages)})
                buffer.clear()
                pages.clear()

        for item in self.items:
            if item.layer != "body":
                continue
            if item.kind in ("title", "heading"):
                flush()
                level = 0 if item.kind == "title" else max(1, item.level)
                while path and path[-1][0] >= level:
                    path.pop()
                path.append((level, item.text))
                continue
            text = markdown_table(item) if item.kind == "table" else item.text
            if not text:
                continue
            if buffer and sum(len(b) for b in buffer) + len(text) > max_chars:
                flush()
            buffer.append(text)
            if item.prov.page is not None:
                pages.add(item.prov.page)
        flush()
        return chunks


def markdown_table(item: Item) -> str:
    """A Markdown table; a merged cell's text appears once, its other positions empty."""
    if not item.cells:
        return ""
    rows = item.rows or max(c.row + c.row_span for c in item.cells)
    cols = item.cols or max(c.col + c.col_span for c in item.cells)
    grid = [["" for _ in range(cols)] for _ in range(rows)]
    for cell in item.cells:
        if cell.row < rows and cell.col < cols:
            grid[cell.row][cell.col] = cell.text.replace("|", "\\|").replace("\n", " ")
    lines = ["| " + " | ".join(grid[0]) + " |", "|" + " --- |" * cols]
    lines += ["| " + " | ".join(r) + " |" for r in grid[1:]]
    return "\n".join(lines)
