"""Convert HTML into Word (.docx), Excel (.xlsx) and PDF files."""

from . import rules
from .converter import Converter, to_docx, to_pdf, to_xlsx
from .model import Document
from .options import DocxOptions, PdfOptions, XlsxOptions

__version__ = "0.1.0"

__all__ = [
    "Converter", "Document", "DocxOptions", "PdfOptions", "XlsxOptions",
    "rules", "to_docx", "to_pdf", "to_xlsx", "__version__",
]
