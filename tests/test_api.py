import asyncio

import pytest

import htmlexport
from htmlexport import DocxOptions, cli


def test_top_level_shortcuts(tmp_path):
    out = tmp_path / "a.docx"
    data = htmlexport.to_docx("<p>x</p>", out)
    assert out.read_bytes() == data and data[:2] == b"PK"


def test_options_object_and_kwargs_are_exclusive(conv):
    with pytest.raises(TypeError):
        conv.to_docx("<p>x</p>", options=DocxOptions(), toc=True)


def test_unknown_option_is_reported(conv):
    with pytest.raises(TypeError, match="tocc"):
        conv.to_docx("<p>x</p>", tocc=True)


def test_bad_input_type(conv):
    with pytest.raises(TypeError):
        conv.to_docx([b"<p>bytes</p>"])


def test_async_variant(conv):
    assert asyncio.run(conv.ato_xlsx("<p>x</p>"))[:2] == b"PK"


def test_cli(tmp_path):
    src = tmp_path / "in.html"
    src.write_text("<p>hello</p>", encoding="utf-8")
    for suffix in (".docx", ".xlsx"):
        out = tmp_path / f"out{suffix}"
        assert cli.main([str(src), "-o", str(out)]) == 0
        assert out.read_bytes()[:2] == b"PK"
