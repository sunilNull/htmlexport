"""Command line: htmlexport report.html -o report.docx"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from . import __version__, rules
from .converter import Converter


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="htmlexport", description="Convert HTML to DOCX, XLSX or PDF.")
    ap.add_argument("inputs", nargs="+", help="HTML files ('-' reads standard input)")
    ap.add_argument("-o", "--output", required=True, help="output file: .docx, .xlsx or .pdf")
    ap.add_argument("--toc", action="store_true", help="DOCX: add a table of contents")
    ap.add_argument("--header", help="DOCX/PDF: page header text")
    ap.add_argument("--page-numbers", action="store_true", help="DOCX/PDF: number the pages")
    ap.add_argument("--layout", choices=["document", "tables"], default="document", help="XLSX layout")
    ap.add_argument("--detect-numbers", action="store_true", help="XLSX: store numbers as numbers")
    ap.add_argument("--page-break-before", metavar="SELECTOR", help="CSS selector: new page before")
    ap.add_argument("--skip", metavar="SELECTOR", help="CSS selector: leave these elements out")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    htmls = [sys.stdin.read() if p == "-" else Path(p).read_text(encoding="utf-8") for p in args.inputs]
    rule_list: list[rules.Rule] = []
    if args.page_break_before:
        rule_list.append(rules.page_break_before(args.page_break_before))
    if args.skip:
        rule_list.append(rules.skip(args.skip))
    conv = Converter(rule_list)

    suffix = Path(args.output).suffix.lower()
    if suffix == ".docx":
        conv.to_docx(htmls, args.output, toc=args.toc, header=args.header, page_numbers=args.page_numbers)
    elif suffix == ".xlsx":
        conv.to_xlsx(htmls, args.output, layout=args.layout, detect_numbers=args.detect_numbers)
    elif suffix == ".pdf":
        conv.to_pdf(htmls, args.output, header=args.header, page_numbers=args.page_numbers)
    else:
        ap.error("output must end in .docx, .xlsx or .pdf")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
