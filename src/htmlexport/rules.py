"""Rules select elements with CSS selectors and change how they are exported.

    conv = Converter(rules=[
        rules.page_break_before(".chapter"),
        rules.toc_headings('[class$="_question_heading"]'),
        rules.border_bottom(".question_description", color="#0096C7"),
        rules.skip(".no-export"),
    ])
"""

from __future__ import annotations

import soupsieve
from bs4 import Tag

from .css import parse_color


class Rule:
    def __init__(self, selector: str) -> None:
        self.selector = selector
        self._compiled = soupsieve.compile(selector)  # raises early on a bad selector

    def matches(self, tag: Tag) -> bool:
        return bool(self._compiled.match(tag))

    def __repr__(self) -> str:
        return f"{type(self).__name__}({self.selector!r})"


class SkipRule(Rule):
    """Leave matching elements out, including everything inside them."""


class PageBreakBeforeRule(Rule):
    """Start a new page before matching elements."""


class TocHeadingsRule(Rule):
    """Only headings matching a TocHeadingsRule go into the Word TOC.

    With no TocHeadingsRule at all, every h1-h6 is a TOC heading.
    """


class BorderBottomRule(Rule):
    """Draw a line under the last paragraph of matching elements."""

    def __init__(self, selector: str, color: str = "#000000") -> None:
        super().__init__(selector)
        self.color = parse_color(color) or "000000"


def skip(selector: str) -> SkipRule:
    return SkipRule(selector)


def page_break_before(selector: str) -> PageBreakBeforeRule:
    return PageBreakBeforeRule(selector)


def toc_headings(selector: str) -> TocHeadingsRule:
    return TocHeadingsRule(selector)


def border_bottom(selector: str, color: str = "#000000") -> BorderBottomRule:
    return BorderBottomRule(selector, color)
