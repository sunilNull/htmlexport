import pytest

from htmlexport.model import Heading, ListBlock, PageBreak, Paragraph, Table, layout_table


def runs(conv, html):
    block = conv.parse(html).blocks[0]
    return [(r.text, r.style.bold, r.style.italic) for r in block.runs]


def test_nested_inline_keeps_text_order_and_styles(conv):
    assert runs(conv, "<p><b>Hello <i>world</i></b></p>") == [
        ("Hello ", True, False), ("world", True, True)]


def test_spaces_between_inline_elements_survive(conv):
    assert runs(conv, "<p>Hello <b>world</b> again</p>") == [
        ("Hello ", False, False), ("world", True, False), (" again", False, False)]


def test_whitespace_collapses_and_trims(conv):
    assert runs(conv, "<p>\n   lots   of\n space  </p>") == [("lots of space", False, False)]


def test_bare_text_is_not_dropped(conv):
    assert runs(conv, "<p>plain text</p>") == [("plain text", False, False)]


def test_lexical_markup(conv):
    html = ('<p dir="ltr"><span style="white-space: pre-wrap;">Keep  </span>'
            '<b><strong style="color: rgb(255, 0, 0);">red</strong></b></p>')
    block = conv.parse(html).blocks[0]
    assert [r.text for r in block.runs] == ["Keep  ", "red"]
    assert block.runs[1].style.color == "FF0000"


def test_empty_paragraph_kept_as_blank_line(conv):
    blocks = conv.parse("<p>a</p><p><br></p><p>b</p>").blocks
    assert [len(b.runs) for b in blocks] == [1, 0, 1]


def test_br_becomes_line_break(conv):
    assert [r.text for r in conv.parse("<p>a<br>b<br></p>").blocks[0].runs] == ["a\nb"]


def test_css_values(conv):
    p = conv.parse('<p style="padding-inline-start: calc(2 * 40px); text-align: center">'
                   '<span style="font-size: 10.5pt; background: rgba(0,0,255,0.5); '
                   'font-weight: 700">x</span></p>').blocks[0]
    assert p.indent == 60.0 and p.align == "center"
    style = p.runs[0].style
    assert (style.font_size, style.background, style.bold) == (10.5, "8080FF", True)


def test_style_with_colon_in_value_does_not_crash(conv):
    p = conv.parse('<p style="background: url(http://x.com/a.png) #fff; color: teal">x</p>').blocks[0]
    assert p.runs[0].style.color == "008080"


def test_links(conv):
    p = conv.parse('<p><a href="https://example.com">site</a> and <a>no href</a></p>').blocks[0]
    assert [(r.text, r.href) for r in p.runs] == [("site", "https://example.com"), (" and no href", None)]


def test_lists(conv):
    lst = conv.parse('<ol type="i" start="4"><li>a<ul><li>b</li></ul></li><li>c</li></ol>').blocks[0]
    assert isinstance(lst, ListBlock) and lst.ordered
    assert (lst.start, lst.number_format, len(lst.items)) == (4, "i", 2)
    assert isinstance(lst.items[0].blocks[1], ListBlock)


def test_table_spans_and_sections(conv):
    t = conv.parse("<table><tfoot><tr><td>foot</td></tr></tfoot><thead><tr><th colspan=2>H</th></tr></thead>"
                   "<tr><td rowspan=2>a</td><td>b</td></tr><tr><td>c</td></tr></table>").blocks[0]
    assert isinstance(t, Table)
    n_rows, n_cols, places = layout_table(t)
    assert (n_rows, n_cols) == (4, 2)
    assert [(p.row, p.col, p.rowspan, p.colspan) for p in places] == [
        (0, 0, 1, 2), (1, 0, 2, 1), (1, 1, 1, 1), (2, 1, 1, 1), (3, 0, 1, 1)]
    assert t.rows[0].cells[0].header


def test_rules(rich_conv):
    html = ('<p>intro</p><div class="report"><h2 class="x_question_heading">Q</h2><h3>H</h3>'
            '<div class="desc">D</div><p class="no-export">secret</p></div>')
    blocks = rich_conv.parse(html).blocks
    assert isinstance(blocks[1], PageBreak)
    assert isinstance(blocks[2], Heading) and blocks[2].in_toc
    assert isinstance(blocks[3], Heading) and not blocks[3].in_toc
    assert isinstance(blocks[4], Paragraph) and blocks[4].border_bottom == "0096C7"
    assert "secret" not in str(blocks)


def test_hidden_content_skipped(conv):
    blocks = conv.parse('<p style="display:none">x</p><p hidden>y</p><script>z</script><p>ok</p>').blocks
    assert len(blocks) == 1 and blocks[0].runs[0].text == "ok"


def test_garbage_input_does_not_crash(conv):
    for html in ["", "   ", "<<<>>>", "<table><tr><td colspan=abc rowspan=-3>x</td></tr></table>",
                 "<ol start='x'><li>a</li></ol>", "<p style=';;;:::'>x</p>", "<b><i>unclosed"]:
        conv.parse(html)


def test_control_characters_are_removed(conv):
    p = conv.parse("<p>a\x01b\x00c\x7f</p>").blocks[0]
    assert p.runs[0].text == "abc"


def test_word_soft_line_break_becomes_line_break(conv):
    assert conv.parse("<p>one\x0btwo</p>").blocks[0].runs[0].text == "one\ntwo"


def test_background_inherit_keeps_parent_color(conv):
    p = conv.parse('<p><mark>a <span style="background: inherit">b</span>'
                   '<span style="background: transparent">c</span></mark></p>').blocks[0]
    assert [(r.text, r.style.background) for r in p.runs] == [("a b", "FFFF00"), ("c", None)]


def test_deep_nesting_gives_clear_error(conv):
    with pytest.raises(ValueError, match="nested too deeply"):
        conv.parse("<div>" * 3000 + "x" + "</div>" * 3000)
