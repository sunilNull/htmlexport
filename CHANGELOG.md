# Changelog

## 0.1.0 (unreleased)

First release.

- One parser turns HTML into a shared document model used by every format.
- DOCX: headings, optional TOC, lists with per-list numbering, tables with
  colspan/rowspan, cell shading, hyperlinks and bookmarks, data-URI images,
  header and page numbers.
- XLSX: `document` layout (content flows down the sheet with fitted row
  heights) and `tables` layout (one sheet per table); optional strict
  number detection.
- PDF: WeasyPrint, with local files and remote URLs blocked by default.
- Rules: `skip`, `page_break_before`, `toc_headings`, `border_bottom`,
  all selected with CSS selectors.
- `ato_*` async methods and an `htmlexport` command-line tool.
