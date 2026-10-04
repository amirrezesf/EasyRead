"""EasyRead converters package."""

from .pptx_to_html import convert as html_convert
from .pptx_to_markdown import convert as md_convert

__all__ = [
    "html_convert",
    "md_convert",
]