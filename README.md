# EasyRead

EasyRead is a desktop Python application for turning scattered educational
sources into structured, AI-ready study material. It extracts text, tables,
images, and embedded audio from common study files, transcribes audio with
Whisper, and creates a purpose-specific Persian prompt that can be pasted into
an external language model.

EasyRead is an extraction and preparation tool. It does **not** call an LLM,
generate a finished video, or publish content by itself.

## What It Does

The main workflow is:

1. Select one or more source files in the desktop GUI.
2. Extract text and media into an output directory.
3. Transcribe audio with OpenAI Whisper when transcription is enabled.
4. Choose a study purpose.
5. Generate a Persian prompt referencing the extracted artifacts.
6. Upload the generated Markdown, transcript, and prompt files to the LLM of
	 your choice.

The project is especially useful for lecture slides with recorded audio,
course PDFs, Word handouts, and classroom recordings.

## Supported Sources

| Input | Processing | Typical output |
| --- | --- | --- |
| `.pptx` | Extract slide text, tables, formatting cues, images, and embedded audio | Markdown, image folder, combined MP3, transcript |
| `.docx` | Extract paragraphs and tables | Markdown |
| `.pdf` | Extract selectable text; render pages without text as images | Markdown and page images |
| `.mp3`, `.wav`, `.m4a`, `.aac`, `.wma`, `.ogg`, `.opus`, `.flac` | Transcribe with Whisper | Plain-text transcript |
| `.mp4` | Treat the media as an audio source for Whisper | Plain-text transcript |
| `.txt`, `.md`, `.html`, `.htm` | Copy the source text into the output directory | Text or Markdown copy |

The GUI accepts multiple sources in one run. Existing files in the selected
output directory may be used when generating a prompt, so use a separate
output directory for each processing job when reproducibility matters.

## Prompt Modes

EasyRead generates three prompt styles in Persian:

- **Study guide**: detailed, step-by-step learning material with definitions,
	relationships, examples, comparisons, and key points.
- **Exam booklet**: comprehensive exam preparation material that keeps
	definitions, classifications, formulas, processes, exceptions, and useful
	examples.
- **Night-before exam handout**: highly compressed, scan-friendly revision
	notes focused on definitions, formulas, differences, steps, and essential
	exceptions.

The shared prompt rules instruct the downstream model to use only the supplied
sources, preserve slide/page references, keep English technical terms, retain
meaningful tables, and avoid inventing missing information.

## Main Desktop Application

The normal entry point is:

```text
python main.py
```

The GUI provides:

- Multi-file source selection
- Optional drag-and-drop support
- Output-folder selection
- Whisper transcription enable/disable
- Whisper model selection
- Persian, English, or automatic language detection
- Background processing so the window remains responsive
- Progress and processing logs
- Session save/load
- Persistent copies of extracted session artifacts
- Purpose-specific prompt generation

The session file is stored below the platform's application-data directory.
On Windows, this is normally `%APPDATA%\\EasyRead\\session.json`, with
extracted artifacts in the adjacent `extracted` directory.

## Standalone Tools

The repository also contains command-line utilities that can be used without
the GUI.

### PowerPoint to Markdown

```text
python pptx_to_markdown.py input.pptx output.md
python pptx_to_markdown.py input.pptx output.md --images-dir output_images
```

This converter preserves or approximates:

- Slide numbering
- Text and bullet hierarchy
- Tables
- Images with Markdown references
- Bold and underline
- Non-default font colors
- Text highlighting
- Grouped shapes, where supported

Images are written to an image directory, which defaults to an `images`
folder beside the Markdown output.

### PowerPoint to HTML

```text
python pptx_to_html.py input.pptx output.html
python pptx_to_html.py input.pptx output.html --images-dir output_images
```

The HTML converter emits a self-contained HTML document plus extracted image
files. It uses semantic HTML such as `<strong>`, `<u>`, `<mark>`, tables, and
slide sections. Its default document direction is right-to-left because the
project's generated study prompts are Persian.

### Extract PowerPoint audio

The audio extractor is primarily used by the main pipeline, but it can also be
run directly:

```text
python extract_audio.py
```

The current `__main__` example uses the hard-coded paths
`eghdamat/1.pptx` and `eghdamat/1.mp3`; change those values before using this
file as a standalone command. The extractor reads the PPTX ZIP/XML structure,
finds audio related to each slide, preserves slide order when possible, and
uses numeric media-file order as a fallback. FFmpeg then combines the clips
into one MP3.

### Legacy PowerPoint text helper

`test.py` contains a small older helper that extracts plain slide text into
`extracted.txt`. It is not an automated test suite and does not provide the
format-preserving behavior of `pptx_to_markdown.py`.

## Architecture

```text
main.py
	-> easyread_gui.py
			 -> source_processing.py
						-> pptx_to_markdown.py
						-> extract_audio.py -> FFmpeg
						-> pypdf / PyMuPDF / python-docx
						-> OpenAI Whisper / PyTorch
			 -> prompt_generation.py

Standalone converters:
	pptx_to_markdown.py -> Markdown + images
	pptx_to_html.py    -> HTML + images
```

### Core modules

- `main.py`: starts the desktop application.
- `easyread_gui.py`: Tkinter interface, session management, progress, and
	user-facing orchestration.
- `source_processing.py`: dispatches each source type and manages
	transcription.
- `pptx_to_markdown.py`: format-aware PowerPoint-to-Markdown conversion.
- `pptx_to_html.py`: PowerPoint-to-HTML conversion.
- `extract_audio.py`: extracts and combines embedded PowerPoint audio.
- `convert_emf_and_wmf.py`: converts EMF/WMF images through LibreOffice and
	PyMuPDF.
- `prompt_generation.py`: builds and writes the three prompt templates.
- `test.py`: legacy plain-text PowerPoint extraction example.

## Requirements

### Python packages

The project currently declares Python `>=3.14` in `pyproject.toml`. The
runtime packages listed in `requirements.txt` are:

- `python-pptx` for PowerPoint parsing
- `python-docx` for Word parsing
- `pypdf` for selectable PDF text
- `PyMuPDF` for rendering PDF pages and converting images
- `openai-whisper` for local speech-to-text
- `torch==2.11.0` for Whisper inference

Tkinter is normally included with desktop Python distributions. Optional GUI
enhancements are detected at runtime:

- `tkinterdnd2` enables drag-and-drop.
- `ttkbootstrap` enables the optional theme.

Install them separately if desired:

```text
python -m pip install tkinterdnd2 ttkbootstrap
```

### External tools

Python packages alone are not sufficient for every feature:

- **FFmpeg** is required for combining embedded PowerPoint audio and converting
	it to MP3. The `ffmpeg` executable must be available on `PATH`.
- **LibreOffice** (`soffice`) is required when a PowerPoint contains EMF or WMF
	images. The converter looks on `PATH` and in common Windows installation
	directories.
- **A suitable compute device** is required for practical Whisper performance.
	The current code requests CUDA explicitly when loading Whisper, so the
	machine must have a compatible NVIDIA/CUDA setup unless the code is adjusted
	to select CPU automatically.

## Installation

Create and activate a virtual environment, then install the dependencies:

### Windows PowerShell

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Install FFmpeg separately and verify it is available:

```powershell
ffmpeg -version
```

Then start EasyRead:

```powershell
python main.py
```

On first use, Whisper downloads the selected model. This can require a large
download and additional disk space. Keep the application open while the model
loads; subsequent runs can reuse the in-process model cache.

## Whisper Models

The GUI exposes these model choices:

`tiny`, `tiny.en`, `base`, `base.en`, `small`, `small.en`, `medium`,
`medium.en`, `large`, and `large-v3-turbo`.

General trade-offs:

| Model family | Speed and memory | Accuracy |
| --- | --- | --- |
| `tiny` / `base` | Lowest requirements and fastest startup | Lowest |
| `small` / `medium` | Balanced, but increasingly expensive | Better |
| `large` / `large-v3-turbo` | Highest requirements and slowest startup | Best or near-best |

The `large-v3-turbo` GUI label maps to Whisper's `turbo` model name internally.
The available language choices are automatic detection, English (`en`), and
Persian/Farsi (`fa`).

## Output Layout

For an input such as `lecture.pptx`, a typical output directory can contain:

```text
lecture_text.md
lecture_images/
	slide1_img1.png
	slide2_img1.jpg
lecture_audio.mp3
lecture_audio_transcript.txt
easyread_prompt_study_guide.txt
```

The exact artifacts depend on the input and enabled options. Text files use
UTF-8 encoding. PDF pages without extractable text are rendered as images and
referenced from the generated Markdown, but they are not OCRed.

## Troubleshooting

### `ffmpeg` is not found

Install FFmpeg and add its `bin` directory to `PATH`. Restart the terminal or
VS Code after changing `PATH`, then run `ffmpeg -version`.

### Whisper fails while loading the model

Check that PyTorch, Whisper, and the selected model are installed. The current
implementation calls `whisper.load_model(..., device="cuda")`; on a CPU-only
machine, change the implementation to choose CUDA only when
`torch.cuda.is_available()` is true and otherwise use `cpu`.

### A scanned PDF contains no useful text

EasyRead renders pages with no selectable text as PNG files, but it does not
perform OCR. Use an OCR tool before importing the PDF, or add an OCR stage to
the pipeline.

### PowerPoint images are out of order

Reading order is inferred from shape coordinates (`top`, then `left`). This
works well for many lecture slides but can be wrong for complex layouts,
columns, diagrams, and free-form designs. Review the generated Markdown or
HTML before using it as the only source for an LLM.

### Extracted images are not understood by the LLM

Images are saved and referenced, but EasyRead does not caption, OCR, or
interpret them. A text-only model will see the reference and filename rather
than the visual content unless you run a separate vision-model step.

## Known Limitations

- The current interface is a desktop Tkinter application, not an Android or
	web application.
- Whisper inference is local and can be slow or memory-intensive.
- The current Whisper loading path assumes CUDA; CPU fallback should be added
	before distributing the application broadly.
- FFmpeg is an external runtime dependency.
- LibreOffice is needed for EMF/WMF conversion.
- Speaker notes are not extracted from PowerPoint files.
- PDF image pages are rendered but not OCRed.
- Extracted images are not automatically captioned or analyzed.
- The generated prompt is saved locally; EasyRead does not send it to an LLM.
- The GUI and most status text are in English, while generated prompts are in
	Persian.
- `pyproject.toml` currently lists fewer dependencies than `requirements.txt`.
	Use `requirements.txt` for the complete current installation, and keep the
	two manifests synchronized when adding or removing packages.
- Automated test coverage is minimal; `test.py` is an example helper rather
	than a test suite.

## Android Status

EasyRead is not directly installable on Android in its current form. Tkinter,
the desktop file dialogs, external FFmpeg invocation, and the CUDA-specific
Whisper path are desktop-oriented.

The processing logic could be reused in a future Android version, but the
mobile product would need a new UI and runtime strategy. The most practical
options are:

1. An Android or web frontend that uploads files to a Python backend running
	 EasyRead, Whisper, and FFmpeg.
2. A fully offline Android app using an Android-compatible Whisper runtime such
	 as `whisper.cpp`, likely with a smaller model.

## Development Notes

Compile-check the Python sources with:

```text
python -m compileall .
```

The project is intentionally modular: extraction and prompt generation can be
used independently of the Tkinter GUI. When extending it, keep filesystem and
UI concerns in the GUI layer and keep source conversion in the processing
modules so a future web or Android frontend can reuse the core.

## License and Project Status

No license file is currently included. Add an explicit license before
redistributing the project.

EasyRead is an early-stage personal/prototype project. It has a useful local
processing pipeline, but it still needs stronger automated tests, unified
dependency metadata, robust CPU fallback for Whisper, and packaging work before
it should be treated as a polished cross-platform product.
