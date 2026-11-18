import io

import pytest

from htmlexport import Converter, rules


@pytest.fixture
def conv():
    return Converter()


@pytest.fixture
def rich_conv():
    return Converter(rules=[
        rules.toc_headings('[class$="_question_heading"]'),
        rules.page_break_before(".report"),
        rules.border_bottom(".desc", "#0096C7"),
        rules.skip(".no-export"),
    ])


def read_docx(data: bytes):
    from docx import Document
    return Document(io.BytesIO(data))


def read_xlsx(data: bytes):
    from openpyxl import load_workbook
    return load_workbook(io.BytesIO(data), rich_text=True)
