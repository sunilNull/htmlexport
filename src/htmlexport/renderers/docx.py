"""Document -> .docx

Word is strict about the order of child elements, so raw XML is inserted at
its schema position instead of appended.
"""

from __future__ import annotations

import base64
import io
import logging
import re
from collections.abc import Sequence
from typing import Any

from docx import Document as new_document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph as DocxParagraph

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
)
from ..options import DocxOptions

log = logging.getLogger("htmlexport")

_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
_HEADING_SIZES = {1: 20, 2: 16, 3: 14, 4: 12, 5: 11, 6: 11}
_NUM_FORMATS = {"1": "decimal", "a": "lowerLetter", "A": "upperLetter",
                "i": "lowerRoman", "I": "upperRoman"}
_BULLETS = ["\u2022", "\u25e6", "\u25aa"]
_SAFE_LINK = re.compile(r"^(https?:|mailto:|tel:|ftp:|#)", re.I)

# Schema successors, used to insert elements in the right position.
_RPR_AFTER_SHD = ("w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
                  "w:eastAsianLayout", "w:specVanish", "w:oMath")
_PPR_AFTER_PBDR = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
                   "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                   "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                   "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                   "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                   "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
_TCPR_AFTER_SHD = ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
                   "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
                   "w:tcPrChange")
_SETTINGS_AFTER_UPDATE = {"hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
                          "rsids", "mathPr", "attachedSchema", "themeFontLang",
                          "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                          "doNotAutoCompressPictures", "forceUpgrade", "captions",
                          "readModeInkLockDown", "smartTagType", "schemaLibrary",
                          "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol",
                          "listSeparator"}


def bookmark_name(anchor: str) -> str:
    """Word bookmark names: start with a letter, [A-Za-z0-9_], at most 40 chars."""
    name = re.sub(r"[^A-Za-z0-9_]", "_", anchor)
    if not name or not name[0].isalpha():
        name = "b_" + name
    return name[:40]


class _Target:
    """Where blocks are written: the document body or a table cell."""

    def __init__(self, obj: Any, fresh: DocxParagraph | None = None, in_cell: bool = False):
        self.obj = obj
        self.fresh = fresh  # an existing empty paragraph to use first (cells start with one)
        self.in_cell = in_cell

    def paragraph(self) -> DocxParagraph:
        if self.fresh is not None:
            p, self.fresh = self.fresh, None
            return p
        return self.obj.add_paragraph()

    def table(self, rows: int, cols: int) -> Any:
        if self.in_cell:
            if self.fresh is not None:  # drop the unused empty paragraph above the table
                self.fresh._p.getparent().remove(self.fresh._p)
                self.fresh = None
            table = self.obj.add_table(rows, cols)
            self.fresh = self.obj.paragraphs[-1]  # python-docx adds one after the table
            return table
        return self.obj.add_table(rows=rows, cols=cols)


class DocxRenderer:
    def __init__(self, options: DocxOptions | None = None) -> None:
        self.o = options or DocxOptions()
        self._bookmark_id = 0
        self._bookmarks: set[str] = set()

    def render(self, documents: Sequence[Document]) -> bytes:
        self.doc = new_document()
        self._setup()
        body = _Target(self.doc)
        if self.o.toc:
            self._toc(body)
        for index, document in enumerate(documents):
            if index > 0 and self.o.page_break_between:
                self._page_break(body)
            self._blocks(document.blocks, body)
        buffer = io.BytesIO()
        self.doc.save(buffer)
        return buffer.getvalue()


    def _setup(self) -> None:
        zoom = self.doc.settings.element.find(qn("w:zoom"))
        if zoom is not None and zoom.get(qn("w:percent")) is None:
            zoom.set(qn("w:percent"), "100")  # python-docx's template omits this required attribute
        normal = self.doc.styles["Normal"]
        normal.font.name = self.o.font
        normal.font.size = Pt(self.o.font_size)
        section = self.doc.sections[0]
        if self.o.page_size == "Letter":
            section.page_width, section.page_height = Inches(8.5), Inches(11)
        else:
            section.page_width, section.page_height = Emu(7560310), Emu(10692130)  # A4
        for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            setattr(section, side, Inches(self.o.margins))
        self._text_width = Emu(int(section.page_width or 0) - int(section.left_margin or 0)
                               - int(section.right_margin or 0))
        if self.o.header:
            section.header.paragraphs[0].text = self.o.header
        if self.o.page_numbers:
            p = section.footer.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _add_field(p, "PAGE", "1")

    def _toc(self, body: _Target) -> None:
        if self.o.toc_title:
            title = body.paragraph()
            run = title.add_run(self.o.toc_title)
            run.bold = True
            run.font.size = Pt(16)
        low, high = self.o.toc_levels
        _add_field(body.paragraph(), f'TOC \\o "{low}-{high}" \\h \\z \\u',
                   "Right-click and choose Update Field to build the table of contents.")
        self._page_break(body)
        if self.o.toc_auto_update:
            settings = self.doc.settings.element
            update = OxmlElement("w:updateFields")
            update.set(qn("w:val"), "true")
            for child in settings:
                if child.tag.split("}")[-1] in _SETTINGS_AFTER_UPDATE:
                    child.addprevious(update)
                    break
            else:
                settings.append(update)


    def _blocks(self, blocks: Sequence[Block], target: _Target) -> None:
        previous: Block | None = None
        for block in blocks:
            if isinstance(previous, Table):
                # Spacing after a table; also stops Word joining two adjacent tables.
                target.paragraph()
            self._block(block, target)
            previous = block

    def _block(self, block: Block, target: _Target, level: int = 0) -> None:
        if isinstance(block, Paragraph):
            self._paragraph(block, target)
        elif isinstance(block, Heading):
            self._heading(block, target)
        elif isinstance(block, ListBlock):
            self._list(block, target, level)
        elif isinstance(block, Table):
            self._table(block, target)
        elif isinstance(block, Image):
            self._image(block, target)
        elif isinstance(block, HorizontalRule):
            p = target.paragraph()
            _bottom_border(p, "A0A0A0")
        elif isinstance(block, PageBreak):
            if not target.in_cell:
                self._page_break(target)

    def _paragraph(self, block: Paragraph, target: _Target,
                   num_id: int | None = None, list_indent: float | None = None) -> DocxParagraph:
        p = target.paragraph()
        if num_id is not None:
            p.style = self.doc.styles["List Paragraph"]
            num_pr = p._p.get_or_add_pPr().get_or_add_numPr()
            num_pr.get_or_add_ilvl().val = 0
            num_pr.get_or_add_numId().val = num_id
        elif list_indent is not None:
            p.paragraph_format.left_indent = Pt(list_indent)
        elif block.indent:
            p.paragraph_format.left_indent = Pt(block.indent)
        self._runs(p, block.runs)
        if block.align:
            p.alignment = _ALIGN[block.align]
        if block.border_bottom:
            _bottom_border(p, block.border_bottom)
        if block.anchor:
            self._bookmark(p, block.anchor)
        return p

    def _heading(self, block: Heading, target: _Target) -> None:
        p = target.paragraph()
        if block.in_toc:
            p.style = self.doc.styles[f"Heading {block.level}"]
            runs = block.runs
        else:
            size = _HEADING_SIZES[block.level]
            runs = [Run(r.text, r.style.with_(bold=True, font_size=r.style.font_size or size), r.href)
                    for r in block.runs]
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(12)
        self._runs(p, runs)
        if block.align:
            p.alignment = _ALIGN[block.align]
        if block.indent:
            p.paragraph_format.left_indent = Pt(block.indent)
        if block.border_bottom:
            _bottom_border(p, block.border_bottom)
        if block.anchor:
            self._bookmark(p, block.anchor)

    def _list(self, lst: ListBlock, target: _Target, level: int) -> None:
        num_id = self._numbering(lst, level)
        text_indent = 36.0 * (level + 1)
        for item in lst.items:
            blocks = list(item.blocks)
            if not blocks or not isinstance(blocks[0], Paragraph):
                blocks.insert(0, Paragraph())
            for index, block in enumerate(blocks):
                if index == 0:
                    self._paragraph(block, target, num_id=num_id)  # type: ignore[arg-type]
                elif isinstance(block, Paragraph):
                    self._paragraph(block, target, list_indent=text_indent)
                elif isinstance(block, ListBlock):
                    self._list(block, target, level + 1)
                else:
                    self._block(block, target, level + 1)

    def _numbering(self, lst: ListBlock, level: int) -> int:
        """Create a numbering definition for one list, so numbering restarts per list."""
        numbering = self.doc.part.numbering_part.element
        abstract_ids = [int(e.get(qn("w:abstractNumId"))) for e in numbering.findall(qn("w:abstractNum"))]
        abstract_id = max(abstract_ids, default=0) + 1

        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "singleLevel")
        abstract.append(multi)
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), "0")
        values = [("w:start", str(max(lst.start, 0)) if lst.ordered else "1"),
                  ("w:numFmt", _NUM_FORMATS.get(lst.number_format, "decimal") if lst.ordered else "bullet"),
                  ("w:lvlText", "%1." if lst.ordered else _BULLETS[level % len(_BULLETS)]),
                  ("w:lvlJc", "left")]
        for tag, value in values:
            el = OxmlElement(tag)
            el.set(qn("w:val"), value)
            lvl.append(el)
        ppr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(720 * (level + 1)))
        ind.set(qn("w:hanging"), "360")
        ppr.append(ind)
        lvl.append(ppr)
        abstract.append(lvl)

        first_num = numbering.find(qn("w:num"))  # abstractNum elements must precede num elements
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            numbering.append(abstract)
        return numbering.add_num(abstract_id).numId

    def _table(self, table: Table, target: _Target) -> None:
        n_rows, n_cols, placements = layout_table(table)
        if not n_rows or not n_cols:
            return
        if table.caption:
            cap = target.paragraph()
            cap.add_run(table.caption).bold = True
        tbl = target.table(n_rows, n_cols)
        tbl.style = self.doc.styles["Table Grid"]
        # table.cell() rebuilds the whole cell list on every call, which makes big
        # tables quadratic. Read the grid once instead (once more after merging).
        cells = tbl._cells
        for pl in placements:
            if pl.rowspan > 1 or pl.colspan > 1:
                last = (pl.row + pl.rowspan - 1) * n_cols + pl.col + pl.colspan - 1
                cells[pl.row * n_cols + pl.col].merge(cells[last])
        if any(pl.rowspan > 1 or pl.colspan > 1 for pl in placements):
            cells = tbl._cells

        header_rows = 0
        for r in range(n_rows):
            starting = [pl for pl in placements if pl.row == r]
            if starting and all(pl.cell.header for pl in starting):
                header_rows += 1
            else:
                break
        for r in range(min(header_rows, n_rows - 1)):  # repeat header rows on every page
            tr_pr = tbl.rows[r]._tr.get_or_add_trPr()
            el = OxmlElement("w:tblHeader")
            el.set(qn("w:val"), "true")
            tr_pr.append(el)

        for pl in placements:
            cell = cells[pl.row * n_cols + pl.col]
            if pl.cell.background:
                _shade(cell._tc.get_or_add_tcPr(), pl.cell.background, _TCPR_AFTER_SHD)
            self._blocks(pl.cell.blocks, _Target(cell, fresh=cell.paragraphs[0], in_cell=True))
            if pl.cell.align:
                for p in cell.paragraphs:
                    if p.alignment is None:
                        p.alignment = _ALIGN[pl.cell.align]

    def _image(self, image: Image, target: _Target) -> None:
        data = _load_image(image.src, self.o.image_loader)
        p = target.paragraph()
        if data is None:
            if image.alt:
                p.add_run(f"[{image.alt}]").italic = True
            return
        try:
            width = Pt(image.width) if image.width else None
            height = Pt(image.height) if image.height and not image.width else None
            shape = p.add_run().add_picture(io.BytesIO(data), width=width, height=height)
            if shape.width > self._text_width:
                ratio = self._text_width / shape.width
                shape.width = Emu(int(shape.width * ratio))
                shape.height = Emu(int(shape.height * ratio))
        except Exception as exc:  # unsupported format (SVG, WebP...) or corrupt data
            log.warning("htmlexport: could not embed image %.60s: %s", image.src, exc)
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            if image.alt:
                p.add_run(f"[{image.alt}]").italic = True


    def _runs(self, p: DocxParagraph, runs: Sequence[Run]) -> None:
        for run in runs:
            if run.href and _SAFE_LINK.match(run.href):
                self._hyperlink(p, run)
            else:
                self._style(p.add_run(run.text), run.style)

    def _hyperlink(self, p: DocxParagraph, run: Run) -> None:
        link = OxmlElement("w:hyperlink")
        href = run.href or ""
        if href.startswith("#"):
            link.set(qn("w:anchor"), bookmark_name(href[1:]))
        else:
            r_id = p.part.relate_to(href, RT.HYPERLINK, is_external=True)
            link.set(qn("r:id"), r_id)
        style = run.style
        if not style.color:
            style = style.with_(color=self.o.link_color, underline=True)
        r = p.add_run(run.text)
        self._style(r, style)
        link.append(r._r)
        p._p.append(link)

    def _style(self, r: Any, s: Style) -> None:
        if s == Style():
            return
        if s.bold:
            r.bold = True
        if s.italic:
            r.italic = True
        if s.underline:
            r.underline = True
        if s.strike:
            r.font.strike = True
        if s.superscript:
            r.font.superscript = True
        elif s.subscript:
            r.font.subscript = True
        if s.color:
            r.font.color.rgb = RGBColor.from_string(s.color)
        if s.font_size:
            r.font.size = Pt(s.font_size)
        if s.font_family:
            r.font.name = s.font_family
            fonts = r._r.get_or_add_rPr().get_or_add_rFonts()
            fonts.set(qn("w:eastAsia"), s.font_family)
            fonts.set(qn("w:cs"), s.font_family)
        if s.background:
            _shade(r._r.get_or_add_rPr(), s.background, _RPR_AFTER_SHD)


    def _page_break(self, target: _Target) -> None:
        target.paragraph().add_run().add_break(WD_BREAK.PAGE)

    def _bookmark(self, p: DocxParagraph, anchor: str) -> None:
        name = bookmark_name(anchor)
        if name in self._bookmarks:  # same id in two documents: links go to the first
            return
        self._bookmarks.add(name)
        self._bookmark_id += 1
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(self._bookmark_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(self._bookmark_id))
        p_pr = p._p.pPr
        if p_pr is not None:
            p_pr.addnext(start)
        else:
            p._p.insert(0, start)
        p._p.append(end)


def _shade(parent: Any, color: str, successors: Sequence[str]) -> None:
    for old in parent.findall(qn("w:shd")):
        parent.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    parent.insert_element_before(shd, *successors)


def _bottom_border(p: DocxParagraph, color: str) -> None:
    p_pr = p._p.get_or_add_pPr()
    border = p_pr.find(qn("w:pBdr"))
    if border is None:
        border = OxmlElement("w:pBdr")
        p_pr.insert_element_before(border, *_PPR_AFTER_PBDR)
    bottom = OxmlElement("w:bottom")
    for key, value in (("w:val", "single"), ("w:sz", "8"), ("w:space", "1"), ("w:color", color)):
        bottom.set(qn(key), value)
    border.append(bottom)


def _add_field(p: DocxParagraph, instruction: str, placeholder: str) -> None:
    def field_char(kind: str) -> None:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        p.add_run()._r.append(el)

    field_char("begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    p.add_run()._r.append(instr)
    field_char("separate")
    p.add_run(placeholder)
    field_char("end")


def _load_image(src: str, loader: Any) -> bytes | None:
    if src.startswith("data:"):
        header, _, payload = src.partition(",")
        try:
            if ";base64" in header:
                return base64.b64decode(payload)
            from urllib.parse import unquote_to_bytes
            return unquote_to_bytes(payload)
        except Exception:
            return None
    if loader is not None:
        try:
            return loader(src)
        except Exception as exc:
            log.warning("htmlexport: image loader failed for %.60s: %s", src, exc)
    return None
