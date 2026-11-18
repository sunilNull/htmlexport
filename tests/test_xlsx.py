import pytest

pytest.importorskip("xlsxwriter")
pytest.importorskip("openpyxl")

from conftest import read_xlsx  # noqa: E402

from htmlexport.renderers.xlsx import parse_number  # noqa: E402

TABLE = ("<table><thead><tr><th>Metric</th><th>2023</th></tr></thead>"
         "<tr><td rowspan=2>CO2</td><td>1,250</td></tr><tr><td>12.5%</td></tr>"
         "<tr><td>Code</td><td>0042</td></tr></table>")


@pytest.mark.parametrize("text,expected", [
    ("1,250", (1250, "#,##0")), ("2023", (2023, "0")), ("12.5%", (0.125, "0.0%")),
    ("$40.00", (40.0, '"$"0.00')), ("-3.5", (-3.5, "0.0")), ("\u20b91,00,000", None),
    ("0042", None), ("1234567890123456", None), ("12a", None), ("$5%", None), ("", None),
])
def test_parse_number(text, expected):
    assert parse_number(text) == expected


def test_tables_layout(conv):
    wb = read_xlsx(conv.to_xlsx(["<p>x</p>" + TABLE, TABLE], layout="tables", detect_numbers=True))
    assert wb.sheetnames == ["Table 1", "Table 2"]
    ws = wb["Table 1"]
    assert ws["B2"].value == 1250 and ws["B3"].value == 0.125 and ws["B4"].value == "0042"
    assert ws["B1"].value == "2023"  # header cells are never converted
    assert "A2:A3" in {str(r) for r in ws.merged_cells.ranges}
    assert ws.freeze_panes == "A2"


def test_numbers_are_text_by_default(conv):
    ws = read_xlsx(conv.to_xlsx(TABLE, layout="tables")).active
    assert ws["B2"].value == "1,250"


def test_tables_layout_falls_back_without_tables(conv):
    wb = read_xlsx(conv.to_xlsx("<p>only text</p>", layout="tables"))
    assert wb.active["A1"].value == "only text"


def test_document_layout(conv):
    long = "word " * 400
    wb = read_xlsx(conv.to_xlsx(f"<h1>Title</h1><p>Hi <b>there</b></p><ul><li>one</li></ul><p>{long}</p>"
                                + TABLE, sheet_names=["My: Report?"]))
    assert wb.sheetnames == ["My_ Report_"]
    ws = wb.active
    assert ws["A1"].value == "Title" and ws["A1"].font.b
    assert str(ws["A2"].value) == "Hi there"  # rich text: two differently styled runs
    assert ws["A3"].value == "\u2022 one"
    assert ws.row_dimensions[4].height > 50  # the long paragraph got a taller row
    assert "A1:H1" in {str(r) for r in ws.merged_cells.ranges}


def test_huge_cell_is_truncated_not_lost(conv):
    ws = read_xlsx(conv.to_xlsx("<p>" + "x" * 40000 + "</p>")).active
    assert len(ws["A1"].value) == 32767


def test_duplicate_sheet_names(conv):
    wb = read_xlsx(conv.to_xlsx(["<p>a</p>", "<p>b</p>"], sheet_names=["Same", "same"]))
    assert wb.sheetnames == ["Same", "same (2)"]


def test_invalid_layout(conv):
    with pytest.raises(ValueError):
        conv.to_xlsx("<p>x</p>", layout="grid")


def test_reserved_sheet_name(conv):
    assert read_xlsx(conv.to_xlsx("<p>x</p>", sheet_names=["History"])).sheetnames == ["History (2)"]
