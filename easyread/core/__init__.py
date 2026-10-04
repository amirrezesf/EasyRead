"""EasyRead core package."""

from easyread.core.source_processing import process_source, transcribe_audio
from easyread.core.prompt_generation import PURPOSES, write_prompt, build_prompt

__all__ = [
    "process_source",
    "transcribe_audio",
    "PURPOSES",
    "write_prompt",
    "build_prompt",
]