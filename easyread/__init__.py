"""EasyRead: Extract and transcribe mixed study sources for LLM prompt generation."""

from easyread.core import (
    process_source,
    transcribe_audio,
    PURPOSES,
    write_prompt,
    build_prompt,
)
from easyread.extractors import (
    extract_audio_to_mp3,
    vector_to_jpg,
)
from easyread.converters import (
    html_convert,
    md_convert,
)
from easyread.cli import cli_main
from easyread.gui import gui_main

__version__ = "0.1.0"

__all__ = [
    # Core
    "process_source",
    "transcribe_audio",
    "PURPOSES",
    "write_prompt",
    "build_prompt",
    # Extractors
    "extract_audio_to_mp3",
    "vector_to_jpg",
    # Converters
    "html_convert",
    "md_convert",
    # Entry points
    "cli_main",
    "gui_main",
]