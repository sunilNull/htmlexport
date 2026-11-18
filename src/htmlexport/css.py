"""Parsing for inline style attributes. Unknown or invalid values are ignored."""

from __future__ import annotations

import re

import tinycss2
from tinycss2 import color3

DEFAULT_FONT_SIZE = 11.0  # pt, used to resolve em/% when nothing else is known
ROOT_FONT_SIZE = 12.0  # pt (16px), used for rem

_LENGTH_RE = re.compile(r"^\s*(-?\d*\.?\d+)\s*(px|pt|em|rem|%|cm|mm|in|pc)?\s*$", re.I)
_CALC_RE = re.compile(
    r"^\s*calc\(\s*(?:(\d*\.?\d+)\s*\*\s*)?([^*()]+?)(?:\s*\*\s*(\d*\.?\d+))?\s*\)\s*$", re.I
)
_FONT_KEYWORDS = {
    "xx-small": 7.0, "x-small": 7.5, "small": 10.0, "medium": 12.0,
    "large": 13.5, "x-large": 18.0, "xx-large": 24.0, "xxx-large": 36.0,
}


def parse_declarations(style: str | None) -> dict[str, str]:
    """Return {property: value} for an inline style string (last one wins)."""
    if not style:
        return {}
    result: dict[str, str] = {}
    try:
        nodes = tinycss2.parse_blocks_contents(style, skip_whitespace=True, skip_comments=True)
    except Exception:  # pragma: no cover
        return {}
    for node in nodes:
        if node.type == "declaration":
            result[node.lower_name] = tinycss2.serialize(node.value).strip()
    return result


def parse_color(value: str | None) -> str | None:
    """CSS color -> "RRGGBB", or None if transparent/invalid.

    Semi-transparent colors are blended onto white, which matches how they
    look on a normal page.
    """
    if not value:
        return None
    # "background" shorthand can hold several tokens; pick the first color.
    for token in _split_top_level(value):
        try:
            rgba = color3.parse_color(token)
        except Exception:
            rgba = None
        if rgba is None or rgba == "currentColor" or not hasattr(rgba, "alpha"):
            continue
        if rgba.alpha <= 0:
            return None
        a = rgba.alpha
        r, g, b = (round(255 * (c * a + (1 - a))) for c in (rgba.red, rgba.green, rgba.blue))
        return f"{r:02X}{g:02X}{b:02X}"
    return None


def parse_length(value: str | None, font_size: float = DEFAULT_FONT_SIZE) -> float | None:
    """CSS length -> points. Supports simple calc() such as ``calc(2 * 40px)``."""
    if not value:
        return None
    m = _CALC_RE.match(value)
    if m:
        factor = float(m.group(1) or 1) * float(m.group(3) or 1)
        inner = parse_length(m.group(2), font_size)
        return None if inner is None else inner * factor
    m = _LENGTH_RE.match(value)
    if not m:
        return None
    number = float(m.group(1))
    unit = (m.group(2) or "px").lower()
    factors = {
        "px": 0.75, "pt": 1.0, "cm": 72 / 2.54, "mm": 72 / 25.4, "in": 72.0, "pc": 12.0,
        "em": font_size, "rem": ROOT_FONT_SIZE, "%": font_size / 100,
    }
    return number * factors[unit]


def parse_font_size(value: str | None, parent: float) -> float | None:
    if not value:
        return None
    value = value.strip().lower()
    if value in _FONT_KEYWORDS:
        return _FONT_KEYWORDS[value]
    if value == "smaller":
        return parent * 0.83
    if value == "larger":
        return parent * 1.2
    size = parse_length(value, parent)
    return size if size and size > 0 else None


def first_font_family(value: str | None) -> str | None:
    if not value:
        return None
    first = value.split(",")[0].strip().strip("'\"")
    generic = {"serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "inherit"}
    return None if not first or first.lower() in generic else first


def _split_top_level(value: str) -> list[str]:
    """Split on whitespace outside parentheses: 'url(a b) #fff' -> ['url(a b)', '#fff']."""
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    for ch in value:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch.isspace() and depth == 0:
            if current:
                parts.append("".join(current))
                current = []
        else:
            current.append(ch)
    if current:
        parts.append("".join(current))
    return parts
