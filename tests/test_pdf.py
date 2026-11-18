import io

import pytest

pytest.importorskip("weasyprint")
pypdf = pytest.importorskip("pypdf")

from htmlexport import Converter, rules  # noqa: E402
from htmlexport.renderers.pdf import prepare_html  # noqa: E402


def pdf_text(data: bytes) -> list[str]:
    return [page.extract_text() for page in pypdf.PdfReader(io.BytesIO(data)).pages]


def test_basic_pdf(conv):
    data = conv.to_pdf("<h1>Hello</h1><p>World</p>")
    assert data.startswith(b"%PDF-")
    assert "Hello" in pdf_text(data)[0]


def test_each_document_starts_a_page(conv):
    assert len(pdf_text(conv.to_pdf(["<p>one</p>", "<p>two</p>"]))) == 2


def test_rules_apply_to_pdf():
    conv = Converter([rules.skip(".secret"), rules.page_break_before(".new")])
    pages = pdf_text(conv.to_pdf('<p>a</p><p class="secret">hidden</p><p class="new">b</p>'))
    assert len(pages) == 2 and "hidden" not in "".join(pages)


def test_local_files_are_blocked(conv, tmp_path, caplog):
    secret = tmp_path / "secret.css"
    secret.write_text("p { color: red }")
    data = conv.to_pdf(f'<link rel="stylesheet" href="file://{secret}"><img src="file:///etc/hostname"><p>x</p>')
    assert data.startswith(b"%PDF-")
    assert "blocked" in caplog.text


def test_allowed_dir_is_readable(tmp_path, caplog):
    (tmp_path / "s.css").write_text("p { color: red }")
    Converter().to_pdf(f'<link rel="stylesheet" href="file://{tmp_path}/s.css"><p>x</p>',
                       allowed_dirs=[str(tmp_path)])
    assert "blocked" not in caplog.text


def test_prepare_html_maps_logical_properties():
    html = prepare_html('<ol type="a" start="3"><li style="padding-inline-start: 40px">x</li></ol>')
    assert "padding-left: 40px" in html and "lower-alpha" in html and "list-item 2" in html
