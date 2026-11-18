import pytest

pytest.importorskip("docx")

from conftest import read_docx  # noqa: E402


def body_text(doc):
    return [p.text for p in doc.paragraphs if p.text]


def test_basic_document(conv):
    doc = read_docx(conv.to_docx("<h1>Title</h1><p>Hello <b>bold</b> world</p>"))
    assert body_text(doc) == ["Title", "Hello bold world"]
    assert doc.paragraphs[0].style.name == "Heading 1"
    assert [r.bold for r in doc.paragraphs[1].runs] == [None, True, None]


def test_toc_and_non_toc_headings(rich_conv):
    doc = read_docx(rich_conv.to_docx('<h2 class="a_question_heading">In</h2><h2>Out</h2>', toc=True))
    styles = {p.text: p.style.name for p in doc.paragraphs if p.text in ("In", "Out")}
    assert styles == {"In": "Heading 2", "Out": "Normal"}
    assert 'TOC \\o "1-3"' in doc.element.xml


def test_hyperlinks(conv):
    doc = read_docx(conv.to_docx('<p id="top">See <a href="https://example.com">site</a>'
                                 ' and <a href="#top">top</a> and <a href="javascript:x()">bad</a></p>'))
    xml = doc.element.xml
    assert xml.count("<w:hyperlink") == 2
    assert 'w:anchor="top"' in xml and 'w:name="top"' in xml
    assert any(r.target_ref == "https://example.com" for r in doc.part.rels.values())


def test_lists_restart_numbering(conv):
    doc = read_docx(conv.to_docx("<ol><li>a</li><li>b</li></ol><ol><li>c</li></ol>"))
    num_ids = [p._p.pPr.numPr.numId.val for p in doc.paragraphs if p.text]
    assert num_ids[0] == num_ids[1] != num_ids[2]


def test_table_merges_and_shading(conv):
    doc = read_docx(conv.to_docx('<table><tr><th colspan="2">H</th></tr>'
                                 '<tr><td style="background:#ff0000">a</td><td>b</td></tr></table>'))
    table = doc.tables[0]
    assert table.cell(0, 0).text == "H" and table.cell(0, 0)._tc is table.cell(0, 1)._tc
    assert 'w:fill="FF0000"' in table.cell(1, 0)._tc.xml


def test_nested_table_and_multi_paragraph_cell(conv):
    doc = read_docx(conv.to_docx("<table><tr><td><p>one</p><p>two</p>"
                                 "<table><tr><td>inner</td></tr></table></td></tr></table>"))
    cell = doc.tables[0].cell(0, 0)
    assert [p.text for p in cell.paragraphs][:2] == ["one", "two"]
    assert cell.tables[0].cell(0, 0).text == "inner"


def test_multiple_documents_get_page_breaks(conv):
    doc = read_docx(conv.to_docx(["<p>one</p>", "<p>two</p>"]))
    assert 'w:type="page"' in doc.element.xml


def test_data_uri_image(conv):
    png = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
    doc = read_docx(conv.to_docx(f'<img src="data:image/png;base64,{png}" alt="dot">'
                                 '<img src="https://example.com/x.png" alt="remote">'))
    assert len(doc.inline_shapes) == 1
    assert "[remote]" in [p.text for p in doc.paragraphs]


def test_control_characters_do_not_crash(conv):
    doc = read_docx(conv.to_docx('<p>bad\x0bline\x01</p><pre>a\x0cb</pre><a href="https://x.com/\x02">l</a>'))
    assert "bad\nline" in [p.text for p in doc.paragraphs]


def test_duplicate_ids_give_unique_bookmarks(conv):
    xml = read_docx(conv.to_docx(['<p id="a">1</p>', '<p id="a">2</p>'])).element.xml
    assert xml.count('w:name="a"') == 1


def test_large_table_is_fast(conv):
    import time

    rows = "".join(f"<tr><td>{i}</td><td>x</td><td>y</td></tr>" for i in range(1000))
    start = time.perf_counter()
    doc = read_docx(conv.to_docx(f"<table><tr><th colspan=3>H</th></tr>{rows}</table>"))
    assert time.perf_counter() - start < 10
    assert doc.tables[0].cell(1000, 0).text == "999"
