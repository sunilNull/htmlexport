"""Document -> .xlsx

Excel doesn't auto-fit row heights for merged cells, so we estimate the height
from the text and set it ourselves. Text taller than the 409pt row limit is
spread over several merged rows.
"""

from __future__ import annotations

import io
import logging
import math
import re
from collections.abc import Sequence
from typing import Any

import xlsxwriter

from ..model import (
    Block,
    Document,
    Heading,
    HorizontalRule,
    Image,
    ListBlock,
    PageBreak,
    Paragraph,
    Run,
    Style,
    Table,
    layout_table,
    plain_text,
)
from ..options import XlsxOptions

log = logging.getLogger("htmlexport")

MAX_ROW_HEIGHT = 409.0
MAX_CELL_CHARS = 32767
_HEADING_SIZES = {1: 16, 2: 14, 3: 13, 4: 12}
_BULLETS = ["\u2022", "\u25e6", "\u25aa"]
_SHEET_BAD = re.compile(r"[\[\]:*?/\\]")
_NUMBER = re.compile(
    r"^(?P<neg>[-\u2212])?(?P<cur>[$\u20ac\u00a3\u00a5\u20b9])?"
    r"(?P<int>\d{1,3}(?:,\d{3})+|\d+)(?:\.(?P<dec>\d+))?(?P<pct>%)?$"
)
_URL = re.compile(r"^(https?:|mailto:)", re.I)


class XlsxRenderer:
    def __init__(self, options: XlsxOptions | None = None) -> None:
        self.o = options or XlsxOptions()

    def render(self, documents: Sequence[Document]) -> bytes:
        buffer = io.BytesIO()
        workbook_options: dict[str, Any] = {"in_memory": True}
        if (self.o.font, self.o.font_size) != ("Calibri", 11):
            workbook_options["default_format_properties"] = {
                "font_name": self.o.font, "font_size": self.o.font_size}
        self.wb = xlsxwriter.Workbook(buffer, workbook_options)
        self._formats: dict[tuple, Any] = {}
        self._names: set[str] = set()

        tables = [t for d in documents for t in _top_level_tables(d.blocks)]
        if self.o.layout == "tables" and tables:
            for index, table in enumerate(tables):
                ws = self.wb.add_worksheet(self._sheet_name(table.caption, f"Table {index + 1}", index))
                self._table_sheet(ws, table)
        else:
            if self.o.layout == "tables":
                log.info("htmlexport: no tables found, using the document layout")
            for index, document in enumerate(documents):
                ws = self.wb.add_worksheet(self._sheet_name(None, f"Sheet{index + 1}", index))
                self._document_sheet(ws, document)
        if not self.wb.worksheets():
            self.wb.add_worksheet()
        self.wb.close()
        return buffer.getvalue()


    def _sheet_name(self, preferred: str | None, fallback: str, index: int) -> str:
        names = self.o.sheet_names
        if names is not None and index < len(names) and names[index]:
            preferred = names[index]
        name = _SHEET_BAD.sub("_", (preferred or fallback).strip()).strip("'")[:31] or fallback
        candidate, n = name, 2
        # Excel reserves "History" and refuses to open a workbook that uses it.
        while candidate.lower() in self._names or candidate.lower() == "history":
            suffix = f" ({n})"
            candidate = name[: 31 - len(suffix)] + suffix
            n += 1
        self._names.add(candidate.lower())
        return candidate


    def _table_sheet(self, ws: Any, table: Table) -> None:
        n_rows, n_cols, placements = layout_table(table)
        widths = [8.0] * n_cols
        for pl in placements:
            if pl.colspan == 1:
                text = plain_text(self._cell_runs(pl.cell.blocks))
                longest = max((len(line) for line in text.split("\n")), default=0)
                widths[pl.col] = max(widths[pl.col], min(60.0, longest * 1.1 + 2))
        for c, w in enumerate(widths):
            ws.set_column(c, c, w)
        header_rows = self._grid(ws, table, 0, 0, widths)
        if self.o.freeze_header and 0 < header_rows < n_rows:
            ws.freeze_panes(header_rows, 0)


    def _document_sheet(self, ws: Any, document: Document) -> None:
        self.ws = ws
        self.row = 0
        self.breaks: list[int] = []
        self.max_cols = max(1, self.o.columns)
        ws.set_column(0, self.max_cols - 1, self.o.column_width)
        for block in document.blocks:
            self._doc_block(block, level=0)
        if self.breaks:
            ws.set_h_pagebreaks(self.breaks)
        ws.fit_to_pages(1, 0)

    def _doc_block(self, block: Block, level: int) -> None:
        base = {"text_wrap": True, "valign": "top"}
        if isinstance(block, Paragraph):
            props = dict(base, align=block.align or "left")
            indent = min(15, round(block.indent / 18))
            if indent:
                props["indent"] = indent
            if block.border_bottom:
                props.update(bottom=1, bottom_color="#" + block.border_bottom)
            self._flow(block.runs, props, indent)
        elif isinstance(block, Heading):
            size = _HEADING_SIZES.get(block.level)
            runs = [Run(r.text, r.style.with_(bold=True, font_size=r.style.font_size or size), r.href)
                    for r in block.runs]
            props = dict(base, align=block.align or "left")
            if block.border_bottom:
                props.update(bottom=1, bottom_color="#" + block.border_bottom)
            self._flow(runs, props, 0)
        elif isinstance(block, ListBlock):
            self._doc_list(block, level)
        elif isinstance(block, Table):
            if block.caption:
                self._flow([Run(block.caption, Style(bold=True))], dict(base), 0)
            n_rows, n_cols, _ = layout_table(block)
            if n_cols > self.max_cols:
                self.ws.set_column(self.max_cols, n_cols - 1, self.o.column_width)
            self._grid(self.ws, block, self.row, 0, [self.o.column_width] * n_cols)
            self.row += n_rows + 1
        elif isinstance(block, Image):
            if block.alt:
                self._flow([Run(f"[{block.alt}]", Style(italic=True))], dict(base), 0)
        elif isinstance(block, HorizontalRule):
            self._flow([], {"bottom": 1, "bottom_color": "#A0A0A0"}, 0)
        elif isinstance(block, PageBreak):
            if self.row:
                self.breaks.append(self.row)

    def _doc_list(self, lst: ListBlock, level: int) -> None:
        for index, item in enumerate(lst.items):
            marker = (_list_number(lst.start + index, lst.number_format) + ". " if lst.ordered
                      else _BULLETS[level % len(_BULLETS)] + " ")
            blocks = item.blocks or [Paragraph()]
            if not isinstance(blocks[0], Paragraph):
                blocks = [Paragraph(), *blocks]
            for j, block in enumerate(blocks):
                if isinstance(block, Paragraph):
                    runs = list(block.runs)
                    if j == 0:
                        style = runs[0].style if runs else Style()
                        runs.insert(0, Run(marker, style.with_(underline=False)))
                    props = {"text_wrap": True, "valign": "top", "align": "left",
                             "indent": min(15, level + 1)}
                    self._flow(runs, props, level + 1)
                elif isinstance(block, ListBlock):
                    self._doc_list(block, level + 1)
                else:
                    self._doc_block(block, level + 1)

    def _flow(self, runs: list[Run], props: dict, indent: int) -> None:
        """Write one block across all document columns, sized to its text."""
        width = self.max_cols * self.o.column_width
        height = self._height(runs, width, indent) if runs else 0.0
        span = max(1, math.ceil(height / MAX_ROW_HEIGHT))
        top, bottom = self.row, self.row + span - 1
        self._write(self.ws, top, 0, bottom, self.max_cols - 1, runs, props)
        if height > self._default_height() + 0.5:
            for r in range(top, bottom + 1):
                self.ws.set_row(r, height / span)
        self.row = bottom + 1


    def _grid(self, ws: Any, table: Table, top: int, left: int, widths: Sequence[float]) -> int:
        """Write a table as a real cell grid. Returns the number of header rows."""
        n_rows, n_cols, placements = layout_table(table)
        heights = [self._default_height()] * n_rows
        spanned: list[tuple[int, int, float]] = []
        header_rows = 0
        for r in range(n_rows):
            starting = [pl for pl in placements if pl.row == r]
            if starting and all(pl.cell.header for pl in starting):
                header_rows += 1
            else:
                break

        for pl in placements:
            runs = self._cell_runs(pl.cell.blocks)
            props: dict[str, Any] = {"text_wrap": True, "valign": "top", "align": pl.cell.align or "left"}
            if self.o.table_borders:
                props["border"] = 1
            if pl.cell.background:
                props["bg_color"] = "#" + pl.cell.background
            elif pl.cell.header:
                props["bg_color"] = "#F2F2F2"
            self._write(ws, top + pl.row, left + pl.col,
                        top + pl.row + pl.rowspan - 1, left + pl.col + pl.colspan - 1,
                        runs, props, numbers=self.o.detect_numbers and not pl.cell.header)
            width = sum(widths[pl.col: pl.col + pl.colspan])
            needed = self._height(runs, width, 0) if runs else 0.0
            if pl.rowspan == 1:
                heights[pl.row] = max(heights[pl.row], needed)
            else:
                spanned.append((pl.row, pl.rowspan, needed))
        for r, span, needed in spanned:
            have = sum(heights[r: r + span])
            if needed > have:
                heights[r + span - 1] += needed - have
        for r, h in enumerate(heights):
            if h > self._default_height() + 0.5:
                ws.set_row(top + r, min(h, MAX_ROW_HEIGHT))
        return header_rows

    def _cell_runs(self, blocks: Sequence[Block], level: int = 0) -> list[Run]:
        """Flatten block content into runs for a single cell."""
        out: list[Run] = []

        def newline() -> None:
            if out:
                out.append(Run("\n"))

        for block in blocks:
            if isinstance(block, Paragraph):
                newline()
                out.extend(block.runs)
            elif isinstance(block, Heading):
                newline()
                out.extend(Run(r.text, r.style.with_(bold=True), r.href) for r in block.runs)
            elif isinstance(block, ListBlock):
                for index, item in enumerate(block.items):
                    marker = (_list_number(block.start + index, block.number_format) + ". "
                              if block.ordered else _BULLETS[level % len(_BULLETS)] + " ")
                    newline()
                    out.append(Run("    " * level + marker))
                    inner = self._cell_runs(item.blocks, level + 1)
                    out.extend(inner)
            elif isinstance(block, Table):
                for row in block.rows:
                    newline()
                    out.append(Run(" | ".join(plain_text(self._cell_runs(c.blocks)).replace("\n", " ")
                                              for c in row.cells)))
            elif isinstance(block, Image) and block.alt:
                newline()
                out.append(Run(f"[{block.alt}]", Style(italic=True)))
        return out


    def _write(self, ws: Any, r1: int, c1: int, r2: int, c2: int, runs: list[Run],
               props: dict, numbers: bool = False) -> None:
        runs = _merge_runs(runs)
        runs = _truncate(runs)
        merged = (r1, c1) != (r2, c2)
        cell_format = self._fmt(props)
        if merged:
            ws.merge_range(r1, c1, r2, c2, "", cell_format)
        if not runs:
            if not merged:
                ws.write_blank(r1, c1, None, cell_format)
            return

        text = plain_text(runs)
        single = len(runs) == 1
        style_props = _font_props(runs[0].style) if single else {}

        if numbers and single:
            number = parse_number(text)
            if number is not None:
                value, num_format = number
                ws.write_number(r1, c1, value, self._fmt({**props, **style_props, "num_format": num_format}))
                return
        if single and runs[0].href and _URL.match(runs[0].href) and len(runs[0].href) <= 2079:
            link_props = {**props, **style_props}
            link_props.setdefault("font_color", "#0563C1")
            link_props.setdefault("underline", 1)
            ws.write_url(r1, c1, runs[0].href, self._fmt(link_props), string=text)
            return
        if single:
            ws.write_string(r1, c1, text, self._fmt({**props, **style_props}))
            return

        fragments: list[Any] = []
        for run in runs:
            font = _font_props(run.style)
            if font:
                fragments.append(self._fmt(font))
            fragments.append(run.text)
        if ws.write_rich_string(r1, c1, *fragments, cell_format) < 0:
            log.warning("htmlexport: rich text rejected at row %d, writing plain text", r1 + 1)
            ws.write_string(r1, c1, text, cell_format)

    def _fmt(self, props: dict) -> Any:
        key = tuple(sorted(props.items()))
        fmt = self._formats.get(key)
        if fmt is None:
            fmt = self._formats[key] = self.wb.add_format(dict(props))
        return fmt

    def _default_height(self) -> float:
        return round(self.o.font_size * 1.36, 1)

    def _height(self, runs: list[Run], width: float, indent: int) -> float:
        """Estimate the row height (pt) wrapped text needs in a cell `width` characters wide."""
        chars_per_line = max(4.0, width - indent * 3)
        lines = sum(max(1, math.ceil(len(line) / chars_per_line))
                    for line in plain_text(runs).split("\n"))
        size = max((r.style.font_size or self.o.font_size for r in runs), default=self.o.font_size)
        return lines * size * 1.36 + (2 if lines > 1 else 0)


def _top_level_tables(blocks: Sequence[Block]) -> list[Table]:
    found: list[Table] = []
    for block in blocks:
        if isinstance(block, Table):
            found.append(block)
        elif isinstance(block, ListBlock):
            for item in block.items:
                found.extend(_top_level_tables(item.blocks))
    return found


def _font_props(s: Style) -> dict:
    props: dict[str, Any] = {}
    if s.bold:
        props["bold"] = True
    if s.italic:
        props["italic"] = True
    if s.underline:
        props["underline"] = 1
    if s.strike:
        props["font_strikeout"] = True
    if s.superscript:
        props["font_script"] = 1
    elif s.subscript:
        props["font_script"] = 2
    if s.color:
        props["font_color"] = "#" + s.color
    if s.font_size:
        props["font_size"] = s.font_size
    if s.font_family:
        props["font_name"] = s.font_family
    return props


def _merge_runs(runs: list[Run]) -> list[Run]:
    out: list[Run] = []
    for run in runs:
        if not run.text:
            continue
        if out and _font_props(out[-1].style) == _font_props(run.style):
            href = out[-1].href if out[-1].href == run.href else None
            out[-1] = Run(out[-1].text + run.text, out[-1].style, href)
        else:
            out.append(run)
    return out


def _truncate(runs: list[Run]) -> list[Run]:
    total = sum(len(r.text) for r in runs)
    if total <= MAX_CELL_CHARS:
        return runs
    log.warning("htmlexport: cell text truncated to Excel's %d character limit", MAX_CELL_CHARS)
    out, budget = [], MAX_CELL_CHARS - 1
    for run in runs:
        if budget <= 0:
            break
        out.append(Run(run.text[:budget], run.style, run.href))
        budget -= len(run.text)
    out[-1] = Run(out[-1].text + "\u2026", out[-1].style, out[-1].href)
    return out


def parse_number(text: str) -> tuple[float, str] | None:
    """Strictly parse a displayed number, keeping how it was written.

    "1,250" -> 1250 with "#,##0"; "2023" -> 2023 with "0" (no separator added);
    "0042" and 16+ digit IDs are left as text.
    """
    m = _NUMBER.match(text.strip())
    if not m or (m["cur"] and m["pct"]):
        return None
    int_part = m["int"].replace(",", "")
    decimals = m["dec"] or ""
    if len(int_part) > 1 and int_part.startswith("0"):
        return None
    if len((int_part + decimals).lstrip("0")) > 15:
        return None
    base = "#,##0" if "," in m["int"] else "0"
    if decimals:
        base += "." + "0" * len(decimals)
    value: float = float(f"{int_part}.{decimals}") if decimals else int(int_part)
    if m["pct"]:
        value, num_format = value / 100, base + "%"
    elif m["cur"]:
        num_format = f'"{m["cur"]}"{base}'
    else:
        num_format = base
    if m["neg"]:
        value = -value
    return value, num_format


def _list_number(n: int, fmt: str) -> str:
    if fmt in ("a", "A") and n > 0:
        letters = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters = chr(97 + rem) + letters
        return letters.upper() if fmt == "A" else letters
    if fmt in ("i", "I") and 0 < n < 4000:
        values = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
                  (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
        roman = ""
        for value, symbol in values:
            while n >= value:
                roman += symbol
                n -= value
        return roman.upper() if fmt == "I" else roman
    return str(n)
