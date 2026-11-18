# htmlexport

Convert HTML into **Word (.docx)**, **Excel (.xlsx)** and **PDF** with one API.

```bash
pip install "htmlexport[all]"      # or pick: [docx], [xlsx], [pdf]
```

```python
from htmlexport import Converter

html = "<h1>Report</h1><p>Revenue grew <b>12%</b> this year.</p>"
conv = Converter()

conv.to_docx(html, "report.docx", toc=True, header="Acme Corp", page_numbers=True)
conv.to_xlsx(html, "report.xlsx")
conv.to_pdf(html, "report.pdf")
```

Every method also returns the file as `bytes`, so a web app can stream it
without writing to disk. Pass a list of HTML strings to combine documents:
each starts a new page (DOCX, PDF) or a new sheet (XLSX).

HTML is parsed once into a small document model that the Word and Excel
writers share, so both get the same content. PDF is rendered by WeasyPrint
directly from the HTML.

## Word (.docx)

```python
conv.to_docx(html, "out.docx",
             toc=True, toc_levels=(1, 2),   # Word builds the TOC from real headings
             header="Acme Corp", page_numbers=True,
             font="Calibri", font_size=11, page_size="A4", margins=1.0)
```

Images in `data:` URIs are embedded. For other sources, pass your own loader,
so you decide what the converter may download:

```python
def load(src: str) -> bytes | None:
    return my_storage.read(src) if src.startswith("s3://my-bucket/") else None

conv.to_docx(html, "out.docx", image_loader=load)
```

## Excel (.xlsx)

Two layouts:

```python
conv.to_xlsx(html, "doc.xlsx")                    # layout="document" (default)
conv.to_xlsx(html, "tables.xlsx", layout="tables")
```

| Layout | Result |
|---|---|
| `document` | Content flows down the sheet: paragraphs and headings span `columns` columns, lists keep their markers, tables become real cell grids. |
| `tables` | Every `<table>` becomes its own sheet, with fitted column widths and a frozen header row. Falls back to `document` if there are no tables. |

Numbers stay text unless you ask for conversion:

```python
conv.to_xlsx(html, "out.xlsx", detect_numbers=True)
```

Detection is deliberately strict: `1,250`, `-3.5`, `12.5%` and `$40.00`
become numbers with a matching format, `2023` stays displayed as `2023`, and
values with leading zeros (`0042`) or more than 15 digits stay text so codes
and IDs are never damaged.

## PDF

```python
conv.to_pdf(html, "out.pdf", page_size="A4", margin="2cm",
            header="Acme Corp", page_numbers=True,
            stylesheets=["h1 { color: #005293; }"])
```

PDF uses [WeasyPrint](https://weasyprint.org/), which needs Pango on the
system ([install guide](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html)).
By default **all external resources are blocked**. Opt in explicitly:

```python
conv.to_pdf(html, "out.pdf",
            allowed_dirs=["/srv/app/static"],  # file:// only inside these folders
            allow_remote=True)                 # http(s) images and stylesheets
```

> Only enable `allow_remote` for HTML you trust. With untrusted HTML it lets
> the document make requests from your server, including to internal addresses.

## Rules

Rules select elements with CSS selectors and apply to all three formats:

```python
from htmlexport import Converter, rules

conv = Converter(rules=[
    rules.page_break_before(".chapter"),                 # new page before each chapter
    rules.toc_headings('[class$="_question_heading"]'),  # only these headings go in the TOC
    rules.border_bottom(".question_description", color="#0096C7"),
    rules.skip(".no-export, .internal-note"),            # leave these out
])
```

## Async (FastAPI, aiohttp, ...)

Conversion is CPU-bound, so the `ato_*` methods run it in a worker thread:

```python
@app.get("/report.docx")
async def report():
    data = await conv.ato_docx(html, toc=True)
    return Response(data, media_type=
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
```

## Command line

```bash
htmlexport report.html -o report.docx --toc --header "Acme Corp"
htmlexport a.html b.html -o combined.pdf --page-numbers
htmlexport data.html -o data.xlsx --layout tables --detect-numbers
```

## Supported HTML

Headings `h1`–`h6`; paragraphs and generic containers; `b strong i em u ins
s strike del sup sub mark code small big a br hr img`; `ul ol li` (with
`type` and `start`); `table thead tbody tfoot tr th td caption` with
`colspan`/`rowspan`; `blockquote pre dl dt dd`. Inline CSS: `color`,
`background(-color)`, `font-weight`, `font-style`, `font-size`,
`font-family`, `text-decoration`, `vertical-align`, `text-align`,
margin/padding indents (including `*-inline-start`), `white-space`,
`display: none` and page-break properties.

Known limits: stylesheets in `<style>` tags are used by the PDF output only;
Excel cannot link part of a cell, so links inside longer text become plain
text there.

## Development

```bash
pip install -e ".[all]" --group dev
pytest
```

## License

MIT
