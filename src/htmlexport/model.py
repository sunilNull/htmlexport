"""Intermediate document model shared by the DOCX and XLSX renderers."""

from __future__ import annotations

from dataclasses import dataclass, field, replace


@dataclass(frozen=True)
class Style:
    """Character formatting. Immutable, so it can be passed down the tree safely."""

    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    superscript: bool = False
    subscript: bool = False
    color: str | None = None  # "RRGGBB"
    background: str | None = None  # "RRGGBB"
    font_size: float | None = None  # points
    font_family: str | None = None

    def with_(self, **changes: object) -> Style:
        return replace(self, **changes)  # type: ignore[arg-type]


DEFAULT_STYLE = Style()


@dataclass
class Run:
    """A piece of text with one style. A text of "\\n" is a line break."""

    text: str
    style: Style = DEFAULT_STYLE
    href: str | None = None


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    align: str | None = None  # left | center | right | justify
    indent: float = 0.0  # points
    border_bottom: str | None = None  # "RRGGBB"
    anchor: str | None = None  # from the element's id attribute


@dataclass
class Heading:
    level: int
    runs: list[Run] = field(default_factory=list)
    align: str | None = None
    indent: float = 0.0
    in_toc: bool = True
    border_bottom: str | None = None
    anchor: str | None = None


@dataclass
class ListItem:
    blocks: list[Block] = field(default_factory=list)


@dataclass
class ListBlock:
    ordered: bool
    items: list[ListItem] = field(default_factory=list)
    start: int = 1
    number_format: str = "1"  # 1 | a | A | i | I


@dataclass
class TableCell:
    blocks: list[Block] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    header: bool = False
    background: str | None = None
    align: str | None = None


@dataclass
class TableRow:
    cells: list[TableCell] = field(default_factory=list)


@dataclass
class Table:
    rows: list[TableRow] = field(default_factory=list)
    caption: str | None = None


@dataclass
class Image:
    src: str
    alt: str = ""
    width: float | None = None  # points
    height: float | None = None  # points


@dataclass
class PageBreak:
    pass


@dataclass
class HorizontalRule:
    pass


Block = Paragraph | Heading | ListBlock | Table | Image | PageBreak | HorizontalRule


@dataclass
class Document:
    blocks: list[Block] = field(default_factory=list)


@dataclass
class CellPlacement:
    cell: TableCell
    row: int
    col: int
    rowspan: int
    colspan: int


def layout_table(table: Table) -> tuple[int, int, list[CellPlacement]]:
    """Resolve colspan/rowspan into absolute grid positions.

    Returns (n_rows, n_cols, placements). Spans that run past the end of the
    table are clipped, as browsers do.
    """
    n_rows = len(table.rows)
    occupied: set[tuple[int, int]] = set()
    placements: list[CellPlacement] = []
    n_cols = 0
    for r, row in enumerate(table.rows):
        c = 0
        for cell in row.cells:
            while (r, c) in occupied:
                c += 1
            rs = max(1, min(cell.rowspan, n_rows - r))
            cs = max(1, cell.colspan)
            for dr in range(rs):
                for dc in range(cs):
                    occupied.add((r + dr, c + dc))
            placements.append(CellPlacement(cell, r, c, rs, cs))
            c += cs
            n_cols = max(n_cols, c)
    for _, c in occupied:
        n_cols = max(n_cols, c + 1)
    return n_rows, n_cols, placements


def plain_text(runs: list[Run]) -> str:
    return "".join(r.text for r in runs)
