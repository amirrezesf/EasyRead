"""Source extraction and Whisper transcription used by the EasyRead GUI."""

from pathlib import Path
import json
import time
import threading
from typing import Optional, List, Any, Callable

from extract_audio import extract_audio_to_mp3


AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".wma", ".ogg", ".opus", ".flac", ".mp4"}
POWERPOINT_EXTENSIONS = {".pptx"}
WORD_EXTENSIONS = {".docx"}
PDF_EXTENSIONS = {".pdf"}
TEXT_EXTENSIONS = {".txt", ".md", ".html", ".htm"}
_WHISPER_MODELS = {}
_WHISPER_MODELS_LOCK = threading.Lock()
_WHISPER_MODEL_ALIASES = {"large-v3-turbo": "turbo"}


def _load_whisper_model(model_name: str, log: Callable[[str], Any] = print) -> Any:
    """Thread-safe Whisper model loading with CPU fallback."""
    with _WHISPER_MODELS_LOCK:
        model = _WHISPER_MODELS.get(model_name)
        if model is not None:
            log(f"Using cached Whisper model: {model_name}")
            return model

    log(f"Loading selected Whisper model: {model_name}")
    try:
        import whisper
    except ImportError as exc:
        raise RuntimeError("Install openai-whisper to transcribe audio with Whisper.") from exc

    # Try CUDA first, fall back to CPU
    try:
        model = whisper.load_model(model_name, device="cuda")
        log(f"Whisper model loaded on CUDA: {model_name}")
    except Exception as e:
        log(f"CUDA not available ({e}), loading on CPU...")
        model = whisper.load_model(model_name, device="cpu")
        log(f"Whisper model loaded on CPU: {model_name}")

    with _WHISPER_MODELS_LOCK:
        _WHISPER_MODELS[model_name] = model
    return model


def _normalize_whisper_language(language: Optional[str]) -> str:
    if language is None:
        return "auto"
    normalized = str(language).strip().lower()
    aliases = {
        "auto": "auto",
        "automatic": "auto",
        "english": "en",
        "en": "en",
        "persian": "fa",
        "farsi": "fa",
        "fa": "fa",
    }
    return aliases.get(normalized, normalized)


def _write_text(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def _extract_ppt_text(source: Path, output_dir: Path, log: Callable[[str], Any] = print) -> Path:
    from pptx_to_markdown import convert

    output = output_dir / f"{source.stem}_text.md"
    images_dir = output_dir / f"{source.stem}_images"
    convert(source, output, images_dir, log=log)
    log(f"Extracted PowerPoint Markdown: {output}")
    return output


def _extract_docx_text(source: Path, output_dir: Path, log: Callable[[str], Any] = print) -> Path:
    from docx import Document

    document = Document(source)
    lines = [f"# {source.stem}", ""]
    lines.extend(p.text.strip() for p in document.paragraphs if p.text.strip())
    for table in document.tables:
        rows = [[cell.text.strip().replace("|", "\\|") for cell in row.cells] for row in table.rows]
        if rows:
            lines.extend(["", "| " + " | ".join(rows[0]) + " |", "| " + " | ".join("---" for _ in rows[0]) + " |"])
            lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    output = _write_text(output_dir / f"{source.stem}_text.md", "\n".join(lines) + "\n")
    log(f"Extracted Word Markdown: {output}")
    return output


def _extract_pdf(source: Path, output_dir: Path, log: Callable[[str], Any] = print) -> List[Path]:
    from pypdf import PdfReader

    reader = PdfReader(str(source))
    pages = [(page.extract_text() or "").strip() for page in reader.pages]
    lines = [f"# {source.stem}", ""]
    for page_number, page in enumerate(pages, start=1):
        if page:
            lines.extend([f"## Page {page_number}", "", page, ""])
    artifacts: List[Path] = []
    pages_without_text = [index for index, page in enumerate(pages) if not page]
    has_some_text = any(pages)
    all_blank = not has_some_text

    if has_some_text:
        output = output_dir / f"{source.stem}_text.md"
        artifacts.append(_write_text(output, "\n".join(lines).rstrip() + "\n"))
        log(f"Extracted PDF Markdown: {output}")

    if pages_without_text:
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError("Some PDF pages have no extractable text; install PyMuPDF to render them as images.") from exc
        image_dir = output_dir / f"{source.stem}_images"
        image_dir.mkdir(parents=True, exist_ok=True)
        document = fitz.open(source)
        for page_number in (index + 1 for index in pages_without_text):
            page = document[page_number - 1]
            image_path = image_dir / f"page{page_number}.png"
            page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).save(image_path)
            artifacts.append(image_path)
            lines.extend([f"## Page {page_number}", "", f"![Page {page_number}]({image_dir.name}/{image_path.name})", ""])
        if all_blank:
            output = output_dir / f"{source.stem}_text.md"
            artifacts.append(_write_text(output, "\n".join(lines).rstrip() + "\n"))
            log(f"Created scanned PDF Markdown with image references: {output}")
        elif has_some_text:
            output.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
        log(f"Rendered {len(pages_without_text)} PDF page image(s) without extractable text: {image_dir}")
    return artifacts


def transcribe_audio(
    audio_path: Path,
    output_path: Path,
    model_name: str = "large-v3-turbo",
    language: str = "fa",
    log: Callable[[str], Any] = print,
    cancel_event: Optional[threading.Event] = None,
) -> Optional[Path]:
    """Transcribe an audio file with OpenAI Whisper and write plain text."""
    whisper_model_name = _WHISPER_MODEL_ALIASES.get(model_name, model_name)
    model = _load_whisper_model(whisper_model_name, log=log)

    whisper_language = _normalize_whisper_language(language)
    log(f"Transcription started: {audio_path.name} (language={whisper_language})")
    started_at = time.perf_counter()
    transcribe_kwargs = {}
    if whisper_language != "auto":
        transcribe_kwargs["language"] = whisper_language

    # Whisper doesn't natively support cancellation, but we can check between segments
    result = model.transcribe(str(audio_path), **transcribe_kwargs)

    # Check for cancellation after transcription
    if cancel_event and cancel_event.is_set():
        log("Transcription cancelled")
        return None

    segments = result.get("segments", [])
    elapsed_seconds = time.perf_counter() - started_at
    lines = [segment.get("text", "").strip() for segment in segments if segment.get("text", "").strip()]
    segment_count = len(segments)
    log(
        f"Transcription finished: {audio_path.name} "
        f"({segment_count} segments, {elapsed_seconds:.1f}s)"
    )
    output = _write_text(Path(output_path), "\n".join(lines))
    log(f"Transcription written: {output}")
    return output


def process_source(
    source_path: Path,
    output_root: Path,
    transcribe: bool = True,
    model_name: str = "large-v3-turbo",
    language: str = "fa",
    log: Callable[[str], Any] = print,
    cancel_event: Optional[threading.Event] = None,
) -> List[Path]:
    """Process one source and return paths created for it."""
    source = Path(source_path)
    output_dir = Path(output_root)
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower()
    artifacts: List[Path] = []
    audio_to_transcribe: List[Path] = []

    if suffix in POWERPOINT_EXTENSIONS:
        artifacts.append(_extract_ppt_text(source, output_dir, log))
        combined_audio = output_dir / f"{source.stem}_audio.mp3"
        try:
            extract_audio_to_mp3(source, str(combined_audio), log=log)
            artifacts.append(combined_audio)
            audio_to_transcribe.append(combined_audio)
        except ValueError as exc:
            log(f"No embedded PowerPoint audio: {exc}")
    elif suffix in AUDIO_EXTENSIONS:
        artifacts.append(source)
        audio_to_transcribe.append(source)
    elif suffix in PDF_EXTENSIONS:
        artifacts.extend(_extract_pdf(source, output_dir, log))
    elif suffix in WORD_EXTENSIONS:
        artifacts.append(_extract_docx_text(source, output_dir, log))
    elif suffix in TEXT_EXTENSIONS:
        # Try UTF-8 first, fall back to common Windows encodings
        text: Optional[str] = None
        for encoding in ("utf-8", "cp1252", "cp1256", "latin-1"):
            try:
                text = source.read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = source.read_text(encoding="utf-8", errors="replace")
            log(f"Warning: {source.name} had encoding issues, used replacement characters")
        output = _write_text(output_dir / f"{source.stem}_text{source.suffix.lower()}", text)
        artifacts.append(output)
        log(f"Copied text source: {output}")
    else:
        raise ValueError(f"Unsupported source type: {source.suffix or 'no extension'}")

    if transcribe:
        for audio_path in audio_to_transcribe:
            # Check for cancellation before each transcription
            if cancel_event and cancel_event.is_set():
                log("Processing cancelled")
                break
            transcript = output_dir / f"{audio_path.stem}_transcript.txt"
            result = transcribe_audio(audio_path, transcript, model_name=model_name, language=language, log=log, cancel_event=cancel_event)
            if result:
                artifacts.append(result)

    # Write manifest for prompt generation
    _write_manifest(source, artifacts, output_dir, log)

    return artifacts


def _write_manifest(
    source: Path,
    artifacts: List[Path],
    output_dir: Path,
    log: Callable[[str], Any] = print,
) -> None:
    """Write a JSON manifest of generated artifacts for prompt generation."""
    manifest_path = output_dir / f"{source.stem}_manifest.json"
    manifest = {
        "source": str(source),
        "source_stem": source.stem,
        "source_suffix": source.suffix,
        "generated_at": time.time(),
        "artifacts": [
            {
                "path": str(artifact),
                "name": Path(artifact).name,
                "type": Path(artifact).suffix.lower(),
            }
            for artifact in artifacts
        ],
    }
    try:
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
        log(f"Manifest written: {manifest_path}")
    except Exception as e:
        log(f"Warning: Failed to write manifest: {e}")