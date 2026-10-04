"""EasyRead extractors package."""

from easyread.extractors.extract_audio import extract_audio_to_mp3
from easyread.extractors.convert_emf_and_wmf import vector_to_jpg

__all__ = [
    "extract_audio_to_mp3",
    "vector_to_jpg",
]