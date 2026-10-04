"""EasyRead converters package."""

from easyread.converters.pptx_to_html import convert as html_convert
from easyread.converters.pptx_to_markdown import convert as md_convert

__all__ = [
    "html_convert",
    "md_convert",
]