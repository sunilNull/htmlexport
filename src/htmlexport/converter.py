from __future__ import annotations

import asyncio
import os
from collections.abc import Sequence
from dataclasses import fields
from typing import Any, TypeVar

from .model import Document
from .options import DocxOptions, PdfOptions, XlsxOptions
from .parser import HTMLParser
from .rules import Rule

HTMLInput = str | Sequence[str]
PathLike = str | os.PathLike[str]
_O = TypeVar("_O", DocxOptions, XlsxOptions, PdfOptions)


class Converter:
    """Convert HTML to DOCX, XLSX or PDF.

    ``html`` can be one string or a list (one section/sheet per item). Options
    are passed as keywords or as an options object, and every method returns
    the file as bytes. ``ato_*`` run in a worker thread for async code.
    """

    def __init__(self, rules: Sequence[Rule] = (), *, html_parser: str = "html.parser") -> None:
        self.rules = list(rules)
        self._parser = HTMLParser(self.rules, html_parser)

    def parse(self, html: str) -> Document:
        """Return the intermediate document model (useful for custom renderers)."""
        try:
            return self._parser.parse(html)
        except RecursionError:
            raise ValueError("HTML is nested too deeply to convert") from None


    def to_docx(self, html: HTMLInput, output: PathLike | None = None, *,
                options: DocxOptions | None = None, **kwargs: Any) -> bytes:
        try:
            from .renderers.docx import DocxRenderer
        except ImportError as exc:
            raise ImportError("DOCX output needs python-docx: pip install 'htmlexport[docx]'") from exc
        opts = _options(DocxOptions, options, kwargs)
        data = DocxRenderer(opts).render([self.parse(h) for h in _as_list(html)])
        return _save(data, output)

    def to_xlsx(self, html: HTMLInput, output: PathLike | None = None, *,
                options: XlsxOptions | None = None, **kwargs: Any) -> bytes:
        try:
            from .renderers.xlsx import XlsxRenderer
        except ImportError as exc:
            raise ImportError("XLSX output needs XlsxWriter: pip install 'htmlexport[xlsx]'") from exc
        opts = _options(XlsxOptions, options, kwargs)
        if opts.layout not in ("document", "tables"):
            raise ValueError("layout must be 'document' or 'tables'")
        data = XlsxRenderer(opts).render([self.parse(h) for h in _as_list(html)])
        return _save(data, output)

    def to_pdf(self, html: HTMLInput, output: PathLike | None = None, *,
               options: PdfOptions | None = None, **kwargs: Any) -> bytes:
        from .renderers.pdf import render_pdf

        opts = _options(PdfOptions, options, kwargs)
        return _save(render_pdf(_as_list(html), opts, self.rules), output)


    async def ato_docx(self, html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
        return await asyncio.to_thread(self.to_docx, html, output, **kwargs)

    async def ato_xlsx(self, html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
        return await asyncio.to_thread(self.to_xlsx, html, output, **kwargs)

    async def ato_pdf(self, html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
        return await asyncio.to_thread(self.to_pdf, html, output, **kwargs)


def _as_list(html: HTMLInput) -> list[str]:
    if isinstance(html, str):
        return [html]
    items = list(html)
    if not all(isinstance(h, str) for h in items):
        raise TypeError("html must be a string or a sequence of strings")
    return items


def _options(cls: type[_O], options: _O | None, kwargs: dict[str, Any]) -> _O:
    if options is not None and kwargs:
        raise TypeError("pass either an options object or keyword options, not both")
    if options is not None:
        return options
    known = {f.name for f in fields(cls)}
    unknown = set(kwargs) - known
    if unknown:
        raise TypeError(f"unknown {cls.__name__} option(s): {', '.join(sorted(unknown))}")
    return cls(**kwargs)


def _save(data: bytes, output: PathLike | None) -> bytes:
    if output is not None:
        with open(output, "wb") as handle:
            handle.write(data)
    return data


_default = Converter()


def to_docx(html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
    """Shortcut for ``Converter().to_docx(...)`` without rules."""
    return _default.to_docx(html, output, **kwargs)


def to_xlsx(html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
    """Shortcut for ``Converter().to_xlsx(...)`` without rules."""
    return _default.to_xlsx(html, output, **kwargs)


def to_pdf(html: HTMLInput, output: PathLike | None = None, **kwargs: Any) -> bytes:
    """Shortcut for ``Converter().to_pdf(...)`` without rules."""
    return _default.to_pdf(html, output, **kwargs)
