"""HTML -> Document.

Each element is turned into a list of items: Runs for inline content, blocks
for everything else. Runs are only grouped into paragraphs when a block ends.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, replace

from bs4 import BeautifulSoup, NavigableString, Tag
from bs4.element import PreformattedString

from . import css
from .model import (
    Block,
    Document,
    Heading,
    HorizontalRule,
    Image,
    ListBlock,
    ListItem,
    PageBreak,
    Paragraph,
    Run,
    Style,
    Table,
    TableCell,
    TableRow,
)
from .rules import BorderBottomRule, PageBreakBeforeRule, Rule, SkipRule, TocHeadingsRule

Item = Run | Block

HEADINGS = {"h1": 1, "h2": 2, "h3": 3, "h4": 4, "h5": 5, "h6": 6}
CONTAINERS = {
    "html", "body", "div", "p", "section", "article", "header", "footer", "main", "nav",
    "aside", "figure", "figcaption", "address", "center", "form", "fieldset", "details",
    "summary", "dl", "dt", "dd", "blockquote", "pre", "li",
}
IGNORED = {
    "head", "script", "style", "title", "meta", "link", "noscript", "template", "iframe",
    "object", "embed", "svg", "canvas", "video", "audio", "input", "button", "select",
    "textarea", "map",
}
BLOCK_DISPLAYS = {"block", "flex", "grid", "list-item", "table", "flow-root"}
TABLE_PARTS = {"table", "thead", "tbody", "tfoot", "tr", "td", "th"}
PAGE_BREAK_VALUES = {"always", "page", "left", "right", "recto", "verso"}
_WS_RE = re.compile(r"[ \t\n\r\f]+")
# Characters XML does not allow; Word and Excel files cannot contain them.
_BAD_CHARS = re.compile(r"[\x00-\x08\x0e-\x1f\x7f\ufffe\uffff\ud800-\udfff]")
_ALIGN = {"left": "left", "start": "left", "right": "right", "end": "right",
          "center": "center", "justify": "justify"}
_LIST_TYPES = {"decimal": "1", "lower-alpha": "a", "lower-latin": "a", "upper-alpha": "A",
               "upper-latin": "A", "lower-roman": "i", "upper-roman": "I"}


class _PreRun(Run):
    """A run whose whitespace must be kept exactly (pre / pre-wrap)."""


@dataclass(frozen=True)
class _Ctx:
    style: Style = Style()
    align: str | None = None
    indent: float = 0.0
    pre: bool = False
    href: str | None = None


class HTMLParser:
    def __init__(self, rules: Sequence[Rule] = (), parser: str = "html.parser") -> None:
        self.parser = parser
        self._skip = [r for r in rules if isinstance(r, SkipRule)]
        self._breaks = [r for r in rules if isinstance(r, PageBreakBeforeRule)]
        self._toc = [r for r in rules if isinstance(r, TocHeadingsRule)]
        self._borders = [r for r in rules if isinstance(r, BorderBottomRule)]


    def parse(self, html: str) -> Document:
        soup = BeautifulSoup(html or "", self.parser)
        root = soup.body or soup
        blocks = self._group(self._walk_children(root, _Ctx()), _Ctx())
        return Document(blocks=_clean_page_breaks(blocks))


    def _walk_children(self, tag: Tag, ctx: _Ctx) -> list[Item]:
        items: list[Item] = []
        for child in tag.children:
            if isinstance(child, PreformattedString):  # comments, CDATA, doctype...
                continue
            if isinstance(child, NavigableString):
                items.extend(self._text(str(child), ctx))
            elif isinstance(child, Tag):
                items.extend(self._element(child, ctx))
        return items

    def _text(self, text: str, ctx: _Ctx) -> list[Item]:
        if "\x0b" in text:  # Word's soft line break
            items: list[Item] = []
            for i, part in enumerate(text.split("\x0b")):
                if i:
                    items.append(Run("\n", ctx.style, ctx.href))
                items.extend(self._text(part, ctx))
            return items
        text = clean_text(text)
        if ctx.pre:
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            return [_PreRun(text, ctx.style, ctx.href)] if text else []
        text = _WS_RE.sub(" ", text)
        return [Run(text, ctx.style, ctx.href)] if text else []

    def _element(self, tag: Tag, ctx: _Ctx) -> list[Item]:
        name = (tag.name or "").lower()
        if name in IGNORED or self._hidden(tag):
            return []
        decl = css.parse_declarations(_attr(tag, "style"))
        if decl.get("display", "").lower() == "none" or decl.get("visibility", "").lower() == "hidden":
            return []

        if name == "br":
            return [Run("\n", ctx.style, ctx.href)]
        if name == "img":
            image = self._image(tag, decl)
            return [image] if image else []

        inner = self._derive(tag, name, decl, ctx)
        if name == "hr":
            blocks: list[Block] = [HorizontalRule()]
        elif name in HEADINGS:
            blocks = [self._heading(tag, HEADINGS[name], inner)]
        elif name in ("ul", "ol"):
            blocks = [self._list(tag, name == "ol", decl, inner)]
        elif name == "table":
            blocks = [self._table(tag, inner)]
        elif name in CONTAINERS or decl.get("display", "").lower() in BLOCK_DISPLAYS:
            blocks = self._container(tag, name, inner)
        else:
            return self._walk_children(tag, inner)
        return list(self._decorate(tag, decl, blocks))

    def _hidden(self, tag: Tag) -> bool:
        return tag.has_attr("hidden") or any(r.matches(tag) for r in self._skip)


    def _derive(self, tag: Tag, name: str, decl: dict[str, str], ctx: _Ctx) -> _Ctx:
        s = ctx.style
        if name in ("b", "strong", "th", "dt"):
            s = s.with_(bold=True)
        elif name in ("i", "em", "cite", "var", "dfn"):
            s = s.with_(italic=True)
        elif name in ("u", "ins"):
            s = s.with_(underline=True)
        elif name in ("s", "strike", "del"):
            s = s.with_(strike=True)
        elif name == "sup":
            s = s.with_(superscript=True, subscript=False)
        elif name == "sub":
            s = s.with_(subscript=True, superscript=False)
        elif name == "mark":
            s = s.with_(background="FFFF00")
        elif name in ("code", "kbd", "samp", "tt", "pre"):
            s = s.with_(font_family="Courier New")
        elif name == "small":
            s = s.with_(font_size=round((s.font_size or css.DEFAULT_FONT_SIZE) * 0.83, 1))
        elif name == "big":
            s = s.with_(font_size=round((s.font_size or css.DEFAULT_FONT_SIZE) * 1.2, 1))
        elif name == "font":
            color = css.parse_color(_attr(tag, "color"))
            face = css.first_font_family(_attr(tag, "face"))
            s = s.with_(color=color or s.color, font_family=face or s.font_family)

        s = _apply_css(s, decl, apply_background=name not in TABLE_PARTS)

        href = ctx.href
        if name == "a" and tag.get("href"):
            href = clean_text(str(tag["href"])).strip() or None

        align = ctx.align
        raw_align = decl.get("text-align") or (tag.get("align") if name not in ("img", "table") else None)
        if raw_align and str(raw_align).lower() in _ALIGN:
            align = _ALIGN[str(raw_align).lower()]
        elif name == "center":
            align = "center"

        indent = ctx.indent
        if name in CONTAINERS or name in ("ul", "ol"):
            for prop in ("margin-left", "padding-left", "margin-inline-start", "padding-inline-start"):
                value = css.parse_length(decl.get(prop), s.font_size or css.DEFAULT_FONT_SIZE)
                if value and value > 0:
                    indent += value
            if name in ("blockquote", "dd"):
                indent += 36.0

        pre = ctx.pre
        ws = decl.get("white-space", "").lower()
        if name == "pre" or ws in ("pre", "pre-wrap", "pre-line", "break-spaces"):
            pre = True
        elif ws in ("normal", "nowrap"):
            pre = False

        return _Ctx(style=s, align=align, indent=indent, pre=pre, href=href)


    def _container(self, tag: Tag, name: str, ctx: _Ctx) -> list[Block]:
        blocks = self._group(self._walk_children(tag, ctx), ctx)
        if not blocks and name == "p":
            # <p></p> or <p><br></p>: editors use these for blank lines.
            blocks = [Paragraph(runs=[], align=ctx.align, indent=ctx.indent)]
        return blocks

    def _heading(self, tag: Tag, level: int, ctx: _Ctx) -> Heading:
        runs: list[Run] = []
        for item in self._walk_children(tag, ctx):
            if isinstance(item, Run):
                runs.append(item)
            elif isinstance(item, (Paragraph, Heading)):
                if runs:
                    runs.append(Run("\n", ctx.style))
                runs.extend(item.runs)
        in_toc = not self._toc or any(r.matches(tag) for r in self._toc)
        return Heading(level=level, runs=_normalize(runs), align=ctx.align,
                       indent=ctx.indent, in_toc=in_toc)

    def _list(self, tag: Tag, ordered: bool, decl: dict[str, str], ctx: _Ctx) -> ListBlock:
        number_format = "1"
        type_attr = tag.get("type")
        if type_attr in ("1", "a", "A", "i", "I"):
            number_format = str(type_attr)
        elif decl.get("list-style-type", "").lower() in _LIST_TYPES:
            number_format = _LIST_TYPES[decl["list-style-type"].lower()]
        try:
            start = int(str(tag.get("start", "1")))
        except ValueError:
            start = 1

        item_ctx = replace(ctx, indent=0.0)  # renderers indent list levels themselves
        lst = ListBlock(ordered=ordered, start=start, number_format=number_format)
        loose: list[Item] = []  # content placed directly in <ul>, outside any <li>

        def add_to_last(items: list[Item]) -> None:
            blocks = self._group(items, item_ctx)
            if not blocks:
                return
            if not lst.items:
                lst.items.append(ListItem())
            lst.items[-1].blocks.extend(blocks)

        for child in tag.children:
            if isinstance(child, Tag) and (child.name or "").lower() == "li":
                add_to_last(loose)
                loose = []
                if self._hidden(child):
                    continue
                li_decl = css.parse_declarations(_attr(child, "style"))
                if li_decl.get("display", "").lower() == "none":
                    continue
                li_ctx = self._derive(child, "li", li_decl, item_ctx)
                li_ctx = replace(li_ctx, indent=0.0)
                lst.items.append(ListItem(blocks=self._group(self._walk_children(child, li_ctx), li_ctx)))
            elif isinstance(child, Tag):
                loose.extend(self._element(child, item_ctx))
            elif isinstance(child, NavigableString) and not isinstance(child, PreformattedString):
                if child.strip():
                    loose.extend(self._text(str(child), item_ctx))
        add_to_last(loose)
        return lst

    def _table(self, tag: Tag, ctx: _Ctx) -> Table:
        head: list[tuple[Tag, bool]] = []
        body: list[tuple[Tag, bool]] = []
        foot: list[tuple[Tag, bool]] = []
        caption: str | None = None
        for child in tag.children:
            if not isinstance(child, Tag):
                continue
            n = (child.name or "").lower()
            if n == "tr":
                body.append((child, False))
            elif n in ("thead", "tbody", "tfoot"):
                target = {"thead": head, "tbody": body, "tfoot": foot}[n]
                target.extend((tr, n == "thead") for tr in child.find_all("tr", recursive=False))
            elif n == "caption":
                caption = _WS_RE.sub(" ", clean_text(child.get_text())).strip() or None

        cell_base = replace(ctx, align=None, indent=0.0)
        table = Table(caption=caption)
        for tr, in_head in head + body + foot:
            if self._hidden(tr):
                continue
            tr_decl = css.parse_declarations(_attr(tr, "style"))
            if tr_decl.get("display", "").lower() == "none":
                continue
            row_bg = css.parse_color(tr_decl.get("background-color") or tr_decl.get("background")
                                     or _attr(tr, "bgcolor"))
            row_ctx = self._derive(tr, "tr", tr_decl, cell_base)
            row = TableRow()
            for cell_tag in tr.find_all(["td", "th"], recursive=False):
                if self._hidden(cell_tag):
                    continue
                cdecl = css.parse_declarations(_attr(cell_tag, "style"))
                cell_ctx = self._derive(cell_tag, cell_tag.name, cdecl, row_ctx)
                if in_head:
                    cell_ctx = replace(cell_ctx, style=cell_ctx.style.with_(bold=True))
                content_ctx = replace(cell_ctx, align=None)
                background = css.parse_color(cdecl.get("background-color") or cdecl.get("background")
                                             or _attr(cell_tag, "bgcolor")) or row_bg
                row.cells.append(TableCell(
                    blocks=self._group(self._walk_children(cell_tag, content_ctx), content_ctx),
                    colspan=_span(cell_tag.get("colspan")),
                    rowspan=_span(cell_tag.get("rowspan")),
                    header=in_head or cell_tag.name == "th",
                    background=background,
                    align=cell_ctx.align,
                ))
            if row.cells:
                table.rows.append(row)
        return table

    def _image(self, tag: Tag, decl: dict[str, str]) -> Image | None:
        src = str(tag.get("src") or "").strip()
        if not src:
            return None

        def size(attr: str) -> float | None:
            value = decl.get(attr) or tag.get(attr)
            if value is None:
                return None
            value = str(value)
            return css.parse_length(value if not value.isdigit() else value + "px")

        alt = clean_text(str(tag.get("alt") or ""))
        return Image(src=src, alt=alt, width=size("width"), height=size("height"))


    def _decorate(self, tag: Tag, decl: dict[str, str], blocks: list[Block]) -> Iterable[Block]:
        if not blocks:
            return blocks
        anchor = tag.get("id")
        if anchor:
            for block in blocks:
                if isinstance(block, (Paragraph, Heading)):
                    block.anchor = block.anchor or str(anchor)
                    break
        for rule in self._borders:
            if rule.matches(tag):
                for block in reversed(blocks):
                    if isinstance(block, (Paragraph, Heading)):
                        block.border_bottom = rule.color
                        break
        before = (decl.get("page-break-before", "").lower() in PAGE_BREAK_VALUES
                  or decl.get("break-before", "").lower() in PAGE_BREAK_VALUES
                  or any(r.matches(tag) for r in self._breaks))
        after = (decl.get("page-break-after", "").lower() in PAGE_BREAK_VALUES
                 or decl.get("break-after", "").lower() in PAGE_BREAK_VALUES)
        result: list[Block] = []
        if before:
            result.append(PageBreak())
        result.extend(blocks)
        if after:
            result.append(PageBreak())
        return result


    def _group(self, items: list[Item], ctx: _Ctx) -> list[Block]:
        blocks: list[Block] = []
        runs: list[Run] = []

        def flush() -> None:
            clean = _normalize(runs)
            if clean:
                blocks.append(Paragraph(runs=clean, align=ctx.align, indent=ctx.indent))
            runs.clear()

        for item in items:
            if isinstance(item, Run):
                runs.append(item)
            else:
                flush()
                blocks.append(item)
        flush()
        return blocks


def _apply_css(s: Style, decl: dict[str, str], apply_background: bool) -> Style:
    if not decl:
        return s
    changes: dict[str, object] = {}
    if "color" in decl:
        color = css.parse_color(decl["color"])
        if color:
            changes["color"] = color
    if apply_background:
        value = decl.get("background-color") or decl.get("background")
        if value:
            color = css.parse_color(value)
            if color:
                changes["background"] = color
            elif value.strip().lower() in ("transparent", "none"):
                changes["background"] = None
    weight = decl.get("font-weight", "").lower()
    if weight:
        if weight in ("bold", "bolder"):
            changes["bold"] = True
        elif weight in ("normal", "lighter"):
            changes["bold"] = False
        elif weight.isdigit():
            changes["bold"] = int(weight) >= 600
    font_style = decl.get("font-style", "").lower()
    if font_style:
        changes["italic"] = font_style in ("italic", "oblique")
    decoration = (decl.get("text-decoration-line") or decl.get("text-decoration") or "").lower()
    if decoration:
        if "none" in decoration.split():
            changes.update(underline=False, strike=False)
        if "underline" in decoration:
            changes["underline"] = True
        if "line-through" in decoration:
            changes["strike"] = True
    valign = decl.get("vertical-align", "").lower()
    if valign == "super":
        changes.update(superscript=True, subscript=False)
    elif valign == "sub":
        changes.update(subscript=True, superscript=False)
    elif valign == "baseline":
        changes.update(subscript=False, superscript=False)
    size = css.parse_font_size(decl.get("font-size"), s.font_size or css.DEFAULT_FONT_SIZE)
    if size:
        changes["font_size"] = round(size, 1)
    family = css.first_font_family(decl.get("font-family"))
    if family:
        changes["font_family"] = family
    return s.with_(**changes) if changes else s


def _normalize(runs: list[Run]) -> list[Run]:
    """Collapse whitespace across run boundaries, trim line ends, merge equal runs."""
    out: list[Run] = []
    prev = "\n"  # the start of a paragraph behaves like the start of a line

    def rstrip_tail() -> None:
        while out and not isinstance(out[-1], _PreRun):
            stripped = out[-1].text.rstrip(" ")
            if stripped:
                out[-1] = Run(stripped, out[-1].style, out[-1].href)
                return
            out.pop()

    for run in runs:
        text = run.text
        if text == "\n" and not isinstance(run, _PreRun):
            rstrip_tail()
            out.append(Run("\n", run.style, run.href))
            prev = "\n"
            continue
        if not isinstance(run, _PreRun) and prev in (" ", "\n"):
            text = text.lstrip(" ")
        if not text:
            continue
        out.append(_PreRun(text, run.style, run.href) if isinstance(run, _PreRun)
                   else Run(text, run.style, run.href))
        prev = text[-1]

    rstrip_tail()
    if out and out[-1].text == "\n":  # a trailing <br> does not create a visible line
        out.pop()
        rstrip_tail()
    if not any(r.text.strip() for r in out) and not any(r.text == "\n" for r in out):
        return []

    merged: list[Run] = []
    for run in out:
        if merged and merged[-1].style == run.style and merged[-1].href == run.href:
            merged[-1] = Run(merged[-1].text + run.text, run.style, run.href)
        else:
            merged.append(Run(run.text, run.style, run.href))
    return merged


def clean_text(text: str) -> str:
    return _BAD_CHARS.sub("", text.replace("\x0b", "\n").replace("\x0c", " "))


def _attr(tag: Tag, name: str) -> str | None:
    """Attribute value as a string (BeautifulSoup returns lists for multi-valued ones)."""
    value = tag.get(name)
    if value is None:
        return None
    return " ".join(value) if isinstance(value, list) else str(value)


def _span(value: object) -> int:
    try:
        return max(1, min(int(str(value)), 1000))
    except (TypeError, ValueError):
        return 1


def _clean_page_breaks(blocks: list[Block]) -> list[Block]:
    """No break at the very start, never two in a row."""
    out: list[Block] = []
    for block in blocks:
        if isinstance(block, PageBreak) and (not out or isinstance(out[-1], PageBreak)):
            continue
        out.append(block)
    while out and isinstance(out[-1], PageBreak):
        out.pop()
    return out
