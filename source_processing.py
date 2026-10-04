"""Source extraction and Whisper transcription used by the EasyRead GUI."""

from pathlib import Path
import time

from extract_audio import extract_audio_to_mp3


AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".wma", ".ogg", ".opus", ".flac", ".mp4"}
POWERPOINT_EXTENSIONS = {".pptx"}
WORD_EXTENSIONS = {".docx"}
PDF_EXTENSIONS = {".pdf"}
TEXT_EXTENSIONS = {".txt", ".md", ".html", ".htm"}
_WHISPER_MODELS = {}
_WHISPER_MODEL_ALIASES = {"large-v3-turbo": "turbo"}


def _normalize_whisper_language(language):
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


def _write_text(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def _extract_ppt_text(source, output_dir, log):
    from pptx_to_markdown import convert

    output = output_dir / f"{source.stem}_text.md"
    images_dir = output_dir / f"{source.stem}_images"
    convert(source, output, images_dir, log=log)
    log(f"Extracted PowerPoint Markdown: {output}")
    return output


def _extract_docx_text(source, output_dir, log):
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


def _extract_pdf(source, output_dir, log):
    from pypdf import PdfReader

    reader = PdfReader(str(source))
    pages = [(page.extract_text() or "").strip() for page in reader.pages]
    lines = [f"# {source.stem}", ""]
    for page_number, page in enumerate(pages, start=1):
        if page:
            lines.extend([f"## Page {page_number}", "", page, ""])
    artifacts = []
    pages_without_text = [index for index, page in enumerate(pages) if not page]
    if any(pages):
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
        if not any(pages):
            output = output_dir / f"{source.stem}_text.md"
            artifacts.append(_write_text(output, "\n".join(lines).rstrip() + "\n"))
            log(f"Created scanned PDF Markdown with image references: {output}")
        elif any(pages):
            output.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
        log(f"Rendered {len(pages_without_text)} PDF page image(s) without extractable text: {image_dir}")
    return artifacts


def transcribe_audio(audio_path, output_path, model_name="large-v3-turbo", language="fa", log=print):
    """Transcribe an audio file with OpenAI Whisper and write plain text."""
    try:
        import whisper
    except ImportError as exc:
        raise RuntimeError("Install openai-whisper to transcribe audio with Whisper.") from exc

    selected_model_name = model_name
    whisper_model_name = _WHISPER_MODEL_ALIASES.get(selected_model_name, selected_model_name)
    model = _WHISPER_MODELS.get(whisper_model_name)
    if model is None:
        log(f"Loading selected Whisper model: {selected_model_name}")
        model = whisper.load_model(whisper_model_name, device="cuda")
        _WHISPER_MODELS[whisper_model_name] = model
        log(f"Whisper model ready: {selected_model_name}")
    else:
        log(f"Using cached Whisper model: {selected_model_name}")

    whisper_language = _normalize_whisper_language(language)
    log(f"Transcription started: {Path(audio_path).name} (language={whisper_language})")
    started_at = time.perf_counter()
    transcribe_kwargs = {}
    if whisper_language != "auto":
        transcribe_kwargs["language"] = whisper_language
    result = model.transcribe(str(audio_path), **transcribe_kwargs)
    segments = result.get("segments", [])
    elapsed_seconds = time.perf_counter() - started_at
    lines = [segment.get("text", "").strip() for segment in segments if segment.get("text", "").strip()]
    segment_count = len(segments)
    log(
        f"Transcription finished: {Path(audio_path).name} "
        f"({segment_count} segments, {elapsed_seconds:.1f}s)"
    )
    output = _write_text(Path(output_path), "\n".join(lines))
    log(f"Transcription written: {output}")
    return output


def process_source(source_path, output_root, transcribe=True, model_name="large-v3-turbo", language="fa", log=print):
    """Process one source and return paths created for it."""
    source = Path(source_path)
    output_dir = Path(output_root)
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower()
    artifacts = []
    audio_to_transcribe = []

    if suffix in POWERPOINT_EXTENSIONS:
        artifacts.append(_extract_ppt_text(source, output_dir, log))
        combined_audio = output_dir / f"{source.stem}_audio.mp3"
        try:
            extract_audio_to_mp3(source, combined_audio, log=log)
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
        output = _write_text(output_dir / f"{source.stem}_text{source.suffix.lower()}", source.read_text(encoding="utf-8"))
        artifacts.append(output)
        log(f"Copied text source: {output}")
    else:
        raise ValueError(f"Unsupported source type: {source.suffix or 'no extension'}")

    if transcribe:
        for audio_path in audio_to_transcribe:
            transcript = output_dir / f"{audio_path.stem}_transcript.txt"
            artifacts.append(transcribe_audio(audio_path, transcript, model_name=model_name, language=language, log=log))
    return artifacts