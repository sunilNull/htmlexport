"""Options for each output format."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from typing import Literal

#: Called with an image ``src``; return the image bytes, or None to skip it.
ImageLoader = Callable[[str], bytes | None]


@dataclass
class DocxOptions:
    toc: bool = False
    toc_title: str | None = "Contents"
    toc_levels: tuple[int, int] = (1, 3)
    #: Ask Word to refresh the TOC on open (Word shows a one-time prompt).
    toc_auto_update: bool = True
    header: str | None = None
    page_numbers: bool = False
    font: str = "Calibri"
    font_size: float = 11
    page_size: Literal["A4", "Letter"] = "A4"
    margins: float = 1.0  # inches
    #: Start every input HTML document on a new page.
    page_break_between: bool = True
    link_color: str = "0563C1"
    #: Loads non-``data:`` images. Default: only ``data:`` URIs are embedded.
    image_loader: ImageLoader | None = None


@dataclass
class XlsxOptions:
    #: "document": content flows down a sheet, tables become real grids.
    #: "tables": every <table> becomes its own sheet.
    layout: Literal["document", "tables"] = "document"
    sheet_names: Sequence[str] | None = None
    #: Document layout: text spans this many columns of ``column_width`` each.
    columns: int = 8
    column_width: float = 15
    font: str = "Calibri"
    font_size: float = 11
    #: Store strictly formatted numbers ("1,250", "12.5%", "$40.00") as numbers.
    detect_numbers: bool = False
    table_borders: bool = True
    freeze_header: bool = True


@dataclass
class PdfOptions:
    page_size: str = "A4"
    margin: str = "2cm"
    header: str | None = None
    page_numbers: bool = False
    #: Extra CSS strings applied to every document.
    stylesheets: Sequence[str] = field(default_factory=tuple)
    base_url: str | None = None
    #: Fetch http(s) resources such as remote images. Off by default:
    #: rendering untrusted HTML with this on allows server-side requests.
    allow_remote: bool = False
    #: Local directories that file:// resources may be read from.
    allowed_dirs: Sequence[str] = field(default_factory=tuple)
    page_break_between: bool = True
