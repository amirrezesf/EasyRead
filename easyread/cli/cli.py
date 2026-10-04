"""EasyRead CLI - Command line interface for extracting and transcribing study sources."""

import argparse
import sys
from pathlib import Path

from easyread.core import process_source
from easyread.core.prompt_generation import PURPOSES, write_prompt


def _parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="EasyRead: Extract and transcribe mixed study sources for LLM prompt generation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  easyread file1.pptx file2.mp3 -o output_dir
  easyread lecture.pptx --transcribe --model large-v3-turbo --language Persian
  easyread --prompt sources.txt output_dir "Exam booklet"
        """,
    )
    ap.add_argument(
        "sources",
        nargs="*",
        type=Path,
        help="Source files to process (pptx, pdf, docx, txt, md, audio/video)",
    )
    ap.add_argument(
        "-o", "--output",
        type=Path,
        default=Path.cwd(),
        help="Output directory (default: current directory)",
    )
    ap.add_argument(
        "--transcribe",
        action="store_true",
        help="Transcribe audio with Whisper",
    )
    ap.add_argument(
        "--no-transcribe",
        dest="transcribe",
        action="store_false",
        help="Skip audio transcription",
    )
    ap.set_defaults(transcribe=True)
    ap.add_argument(
        "--model",
        default="large-v3-turbo",
        choices=["tiny", "tiny.en", "base", "base.en", "small", "small.en", "medium", "medium.en", "large", "large-v3-turbo"],
        help="Whisper model size (default: large-v3-turbo)",
    )
    ap.add_argument(
        "--language",
        default="Persian",
        choices=["Automatic", "Persian", "English"],
        help="Transcription language (default: Persian)",
    )
    ap.add_argument(
        "--prompt",
        nargs=2,
        metavar=("OUTPUT_DIR", "PURPOSE"),
        help="Generate prompt from already-processed sources. PURPOSE: Study guide | Exam booklet | Night-before exam handout",
    )
    ap.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Verbose output",
    )
    return ap.parse_args()


def _log_verbose(message: str, verbose: bool) -> None:
    if verbose:
        print(message, file=sys.stderr)


def main() -> int:
    args = _parse_args()

    # Prompt generation mode
    if args.prompt:
        output_dir, purpose = args.prompt
        if purpose not in PURPOSES:
            print(f"Error: Unknown prompt purpose: {purpose}", file=sys.stderr)
            print(f"Valid purposes: {', '.join(PURPOSES.keys())}", file=sys.stderr)
            return 1
        try:
            # Need to find sources from output_dir - they must have been processed before
            output_path = Path(output_dir)
            # Find all source stems by looking for manifest files
            sources = []
            for manifest in output_path.glob("*_manifest.json"):
                try:
                    import json
                    with open(manifest, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    # The source in manifest is the original source path
                    source_path = Path(data.get("source", ""))
                    if source_path.exists():
                        sources.append(source_path)
                except Exception:
                    pass
            if not sources:
                print("Error: No processed sources found in output directory", file=sys.stderr)
                return 1
            prompt_path = write_prompt(sources, output_path, purpose)
            print(f"Prompt created: {prompt_path}")
            return 0
        except Exception as e:
            print(f"Error generating prompt: {e}", file=sys.stderr)
            return 1

    # Source processing mode
    if not args.sources:
        print("Error: No source files specified", file=sys.stderr)
        return 1

    output_dir = args.output
    output_dir.mkdir(parents=True, exist_ok=True)

    errors = []
    for source in args.sources:
        if not source.exists():
            print(f"Error: Source not found: {source}", file=sys.stderr)
            errors.append(str(source))
            continue

        _log_verbose(f"Processing {source}...", args.verbose)
        try:
            artifacts = process_source(
                source,
                output_dir,
                args.transcribe,
                args.model,
                language=args.language,
                log=lambda msg: _log_verbose(msg, args.verbose),
            )
            for artifact in artifacts:
                _log_verbose(f"  Created: {artifact}", args.verbose)
        except Exception as e:
            print(f"Error processing {source}: {e}", file=sys.stderr)
            errors.append(str(source))

    if errors:
        print(f"\nFinished with {len(errors)} error(s)", file=sys.stderr)
        return 1

    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())