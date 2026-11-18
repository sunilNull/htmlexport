"""HTML -> PDF via WeasyPrint.

External resources are blocked unless explicitly allowed (see PdfOptions).
"""

from __future__ import annotations

import logging
import os
from collections.abc import Sequence
from typing import Any
from urllib.parse import urlsplit
from urllib.request import url2pathname

from bs4 import BeautifulSoup

from ..css import parse_declarations
from ..options import PdfOptions
from ..rules import BorderBottomRule, PageBreakBeforeRule, Rule, SkipRule

log = logging.getLogger("htmlexport")

# WeasyPrint does not support these logical properties; map them for LTR text.
_LOGICAL = {"padding-inline-start": "padding-left", "padding-inline-end": "padding-right",
            "margin-inline-start": "margin-left", "margin-inline-end": "margin-right"}


_OL_TYPES = {"1": "decimal", "a": "lower-alpha", "A": "upper-alpha",
             "i": "lower-roman", "I": "upper-roman"}


class BlockedResource(ValueError):
    pass


def render_pdf(htmls: Sequence[str], options: PdfOptions | None = None,
               rules: Sequence[Rule] = ()) -> bytes:
    try:
        import weasyprint
    except ImportError as exc:  # pragma: no cover
        raise ImportError("PDF output needs WeasyPrint: pip install 'htmlexport[pdf]'") from exc

    o = options or PdfOptions()
    fetcher = _make_fetcher(weasyprint, o)
    stylesheets = [weasyprint.CSS(string=_page_css(o), url_fetcher=fetcher)]
    stylesheets += [weasyprint.CSS(string=s, url_fetcher=fetcher) for s in o.stylesheets]

    documents = [
        weasyprint.HTML(string=prepare_html(h, rules), base_url=o.base_url, url_fetcher=fetcher)
        .render(stylesheets=stylesheets)
        for h in htmls
    ]
    if not documents:
        documents = [weasyprint.HTML(string="").render(stylesheets=stylesheets)]
    if o.page_break_between or len(documents) == 1:
        pages = [page for document in documents for page in document.pages]
        return documents[0].copy(pages).write_pdf()
    combined = "".join(prepare_html(h, rules) for h in htmls)
    return weasyprint.HTML(string=combined, base_url=o.base_url, url_fetcher=fetcher).write_pdf(
        stylesheets=stylesheets)


def prepare_html(html: str, rules: Sequence[Rule] = ()) -> str:
    """Apply rules and fix CSS WeasyPrint does not understand."""
    soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup.find_all(True):
        if tag.decomposed:
            continue
        if any(isinstance(r, SkipRule) and r.matches(tag) for r in rules):
            tag.decompose()
            continue
        extra = []
        if any(isinstance(r, PageBreakBeforeRule) and r.matches(tag) for r in rules):
            extra.append("break-before: page")
        for r in rules:
            if isinstance(r, BorderBottomRule) and r.matches(tag):
                extra.append(f"border-bottom: 1px solid #{r.color}")
        if tag.name == "ol":
            # WeasyPrint ignores these HTML attributes; express them as CSS.
            list_type = _OL_TYPES.get(str(tag.get("type", "")))
            if list_type:
                extra.append(f"list-style-type: {list_type}")
            start = str(tag.get("start", "")).strip()
            if start.lstrip("-").isdigit():
                extra.append(f"counter-reset: list-item {int(start) - 1}")
        raw = tag.get("style")
        style = " ".join(raw) if isinstance(raw, list) else raw
        if style or extra:
            decl = parse_declarations(style)
            parts = [f"{_LOGICAL.get(k, k)}: {v}" for k, v in decl.items()] + extra
            tag["style"] = "; ".join(parts)
    return str(soup)


def _page_css(o: PdfOptions) -> str:
    margin_boxes = ""
    if o.header:
        margin_boxes += f'@top-left {{ content: "{_css_string(o.header)}"; font-size: 9pt; color: #555; }}'
    if o.page_numbers:
        margin_boxes += '@bottom-right { content: counter(page) " / " counter(pages); font-size: 9pt; }'
    return (f"@page {{ size: {o.page_size}; margin: {o.margin}; {margin_boxes} }}"
            "body { font-family: Calibri, Carlito, 'Liberation Sans', Arial, sans-serif; font-size: 11pt; }"
            "table { border-collapse: collapse; }")


def _css_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def _allowed_file(url: str, allowed_dirs: Sequence[str]) -> bool:
    path = os.path.realpath(url2pathname(urlsplit(url).path))
    for directory in allowed_dirs:
        root = os.path.realpath(directory)
        try:
            if os.path.commonpath([root, path]) == root:
                return True
        except ValueError:  # different drives on Windows
            continue
    return False


def _check(url: str, o: PdfOptions) -> None:
    scheme = urlsplit(url).scheme.lower()
    if scheme == "data":
        return
    if scheme in ("http", "https") and o.allow_remote:
        return
    if scheme == "file" and _allowed_file(url, o.allowed_dirs):
        return
    raise BlockedResource(f"htmlexport blocked resource: {url[:120]}")


def _make_fetcher(weasyprint: Any, o: PdfOptions) -> Any:
    urls = weasyprint.urls
    if hasattr(urls, "URLFetcher"):  # WeasyPrint >= 66

        class SafeFetcher(urls.URLFetcher):  # type: ignore[misc, name-defined]
            def fetch(self, url: str, headers: Any = None) -> Any:
                _check(url, o)
                return super().fetch(url, headers)

        return SafeFetcher(allow_redirects=False)

    def fetch(url: str, *args: Any, **kwargs: Any) -> Any:  # older WeasyPrint
        _check(url, o)
        return urls.default_url_fetcher(url, *args, **kwargs)

    return fetch
