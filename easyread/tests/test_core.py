"""Basic tests for EasyRead core modules."""

import tempfile
from pathlib import Path

import pytest


def test_normalize_whisper_language():
    from easyread.core.source_processing import _normalize_whisper_language

    assert _normalize_whisper_language(None) == "auto"
    assert _normalize_whisper_language("auto") == "auto"
    assert _normalize_whisper_language("Automatic") == "auto"
    assert _normalize_whisper_language("english") == "en"
    assert _normalize_whisper_language("en") == "en"
    assert _normalize_whisper_language("persian") == "fa"
    assert _normalize_whisper_language("farsi") == "fa"
    assert _normalize_whisper_language("fa") == "fa"
    assert _normalize_whisper_language("unknown") == "unknown"


def test_safe_name():
    from easyread.core.prompt_generation import _safe_name

    assert _safe_name("Study guide") == "study_guide"
    assert _safe_name("Night-before exam handout") == "night-before_exam_handout"
    assert _safe_name("Exam booklet") == "exam_booklet"


def test_source_processing_imports():
    """Verify all source_processing functions are importable."""
    from easyread.core.source_processing import (
        process_source,
        transcribe_audio,
        _load_whisper_model,
        _normalize_whisper_language,
        _extract_pdf,
        _extract_docx_text,
        _extract_ppt_text,
    )
    assert callable(process_source)
    assert callable(transcribe_audio)
    assert callable(_load_whisper_model)
    assert callable(_normalize_whisper_language)
    assert callable(_extract_pdf)
    assert callable(_extract_docx_text)
    assert callable(_extract_ppt_text)


def test_prompt_generation_imports():
    """Verify all prompt_generation functions are importable."""
    from easyread.core.prompt_generation import (
        build_prompt,
        write_prompt,
        PURPOSES,
        _safe_name,
        _source_artifacts,
        _collect_artifact_names,
    )
    assert callable(build_prompt)
    assert callable(write_prompt)
    assert isinstance(PURPOSES, dict)
    assert "Study guide" in PURPOSES
    assert "Exam booklet" in PURPOSES
    assert "Night-before exam handout" in PURPOSES
    assert callable(_safe_name)
    assert callable(_source_artifacts)
    assert callable(_collect_artifact_names)


def test_pptx_converters_imports():
    """Verify PPTX converters are importable."""
    from easyread.converters.pptx_to_html import convert as html_convert
    from easyread.converters.pptx_to_markdown import convert as md_convert
    assert callable(html_convert)
    assert callable(md_convert)


def test_extract_audio_imports():
    """Verify extract_audio functions are importable."""
    from easyread.extractors.extract_audio import extract_audio_to_mp3
    assert callable(extract_audio_to_mp3)


def test_convert_emf_imports():
    """Verify convert_emf_and_wmf functions are importable."""
    from easyread.extractors.convert_emf_and_wmf import vector_to_jpg
    assert callable(vector_to_jpg)


def test_manifest_writing(tmp_path):
    """Test that _write_manifest creates valid JSON."""
    from easyread.core.source_processing import _write_manifest

    source = tmp_path / "test.pptx"
    source.write_text("dummy")

    artifacts = [tmp_path / "test_text.md", tmp_path / "test_audio.mp3"]
    for a in artifacts:
        a.write_text("dummy")

    _write_manifest(source, artifacts, tmp_path)

    manifest_path = tmp_path / "test_manifest.json"
    assert manifest_path.exists()

    import json
    with open(manifest_path) as f:
        manifest = json.load(f)

    assert manifest["source_stem"] == "test"
    assert manifest["source_suffix"] == ".pptx"
    assert len(manifest["artifacts"]) == 2
    assert manifest["artifacts"][0]["name"] == "test_text.md"
    assert manifest["artifacts"][1]["name"] == "test_audio.mp3"


def test_prompt_purpose_structure():
    """Test that all purposes have required fields."""
    from easyread.core.prompt_generation import PURPOSES

    required_fields = {"role", "goal", "content_rules", "organization"}
    for purpose_name, config in PURPOSES.items():
        assert isinstance(config, dict), f"{purpose_name}: config must be dict"
        assert set(config.keys()) == required_fields, f"{purpose_name}: missing fields"
        for field in required_fields:
            assert isinstance(config[field], str), f"{purpose_name}.{field}: must be string"
            assert len(config[field]) > 0, f"{purpose_name}.{field}: must not be empty"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])