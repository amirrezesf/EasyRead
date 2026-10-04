import zipfile
import subprocess
import tempfile
import re
import shutil
from pathlib import Path
from xml.etree import ElementTree as ET
from typing import List, Optional, Callable, Any


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

AUDIO_EXTENSIONS = {
    ".mp3", ".wav", ".m4a", ".aac",
    ".wma", ".ogg", ".opus", ".flac"
}


def _get_ordered_slide_paths(ppt_zip: zipfile.ZipFile) -> List[str]:
    """Return list of 'ppt/slides/slideN.xml' paths in actual presentation order."""
    pres_xml = ET.fromstring(ppt_zip.read("ppt/presentation.xml"))
    pres_rels = ET.fromstring(ppt_zip.read("ppt/_rels/presentation.xml.rels"))

    rid_to_target = {}
    for rel in pres_rels.findall("rel:Relationship", NS):
        rid_to_target[rel.get("Id")] = rel.get("Target")

    slide_paths = []
    sldIdLst = pres_xml.find("p:sldIdLst", NS)
    for sldId in sldIdLst.findall("p:sldId", NS):
        rid = sldId.get(f"{{{NS['r']}}}id")
        target = rid_to_target.get(rid)
        if target is None:
            continue
        # Targets are relative to ppt/, normalize
        target = target.replace("\\", "/")
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = f"ppt/{target}"
        slide_paths.append(path)

    return slide_paths


def _get_slide_audio_media(ppt_zip: zipfile.ZipFile, slide_path: str) -> List[str]:
    """Return list of media paths (ppt/media/...) referenced by a slide, in the
    order they appear in the slide XML (covers timeline/click order better than
    rels file order alone)."""
    slide_dir = Path(slide_path).parent
    slide_name = Path(slide_path).name
    rels_path = f"{slide_dir}/_rels/{slide_name}.rels"

    try:
        rels_xml = ET.fromstring(ppt_zip.read(rels_path))
    except KeyError:
        return []

    rid_to_media = {}
    for rel in rels_xml.findall("rel:Relationship", NS):
        target = rel.get("Target", "").replace("\\", "/")

        if Path(target).suffix.lower() not in AUDIO_EXTENSIONS:
            continue

        # Targets are typically relative to the slide's directory, e.g.
        # "../media/media1.mp3", but can also appear as "media/media1.mp3"
        # or (rarely) an absolute in-package path like "/ppt/media/media1.mp3".
        # Normalize all of these to "ppt/media/<filename>".
        norm = target.lstrip("/")
        if norm.startswith("ppt/media/"):
            media_path = norm
        elif "media/" in norm:
            # take from the last "media/" occurrence onward
            media_path = "ppt/" + norm[norm.rfind("media/"):]
        else:
            # fallback: just filename, assume it lives in ppt/media/
            media_path = f"ppt/media/{Path(norm).name}"

        rid_to_media[rel.get("Id")] = media_path

    if not rid_to_media:
        return []

    # Walk the slide XML in document order, collecting r:embed / r:link refs
    # that point at audio media, so multiple sounds on one slide keep their
    # in-slide order too.
    slide_xml_raw = ppt_zip.read(slide_path)
    ordered_rids = re.findall(r'r:embed="(rId\d+)"', slide_xml_raw.decode("utf-8", "ignore"))
    ordered_rids += re.findall(r'r:link="(rId\d+)"', slide_xml_raw.decode("utf-8", "ignore"))

    media_in_order = []
    seen = set()
    for rid in ordered_rids:
        media = rid_to_media.get(rid)
        if media and media not in seen:
            media_in_order.append(media)
            seen.add(media)

    # Fallback: if regex approach found nothing but rels did, use rels order
    if not media_in_order:
        media_in_order = list(rid_to_media.values())

    return media_in_order


def _numeric_key(path: Path) -> float:
    """Extract the trailing integer from a filename like 'media23.m4a' -> 23."""
    match = re.search(r"(\d+)", path.stem)
    return int(match.group(1)) if match else float("inf")


def extract_audio_to_mp3(
    powerpoint_file: Path,
    output_file: Path = Path("combined_audio.mp3"),
    log: Callable[[str], Any] = print,
) -> Path:
    powerpoint_file = Path(powerpoint_file)
    output_file = Path(output_file)

    if not powerpoint_file.exists():
        raise FileNotFoundError(f"PowerPoint file not found: {powerpoint_file}")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)

        with zipfile.ZipFile(powerpoint_file, "r") as ppt_zip:
            all_media = [n for n in ppt_zip.namelist() if n.startswith("ppt/media/")]
            if not all_media:
                raise ValueError("No media files were found in the PowerPoint.")

            audio_media_all = [
                m for m in all_media if Path(m).suffix.lower() in AUDIO_EXTENSIONS
            ]
            if not audio_media_all:
                raise ValueError("No audio files were found in the PowerPoint.")

            slide_paths = _get_ordered_slide_paths(ppt_zip)

            # Try the "proper" slide-relationship approach first.
            ordered_audio_media: List[str] = []
            seen = set()
            for slide_path in slide_paths:
                for media in _get_slide_audio_media(ppt_zip, slide_path):
                    if media not in seen:
                        ordered_audio_media.append(media)
                        seen.add(media)

            if not ordered_audio_media:
                # No slide rels exist at all (seen in the wild with some
                # export tools). Fall back to matching mediaN.ext -> slide N
                # numerically, which is correct when the counts line up.
                sorted_media = sorted(audio_media_all, key=lambda m: _numeric_key(Path(m)))

                if len(sorted_media) == len(slide_paths):
                    log(
                        f"No slide relationships found; falling back to numeric "
                        f"filename order (media count {len(sorted_media)} matches "
                        f"slide count {len(slide_paths)})."
                    )
                else:
                    log(
                        f"WARNING: No slide relationships found, and audio file count "
                        f"({len(sorted_media)}) does not match slide count "
                        f"({len(slide_paths)}). Falling back to numeric filename order "
                        f"anyway, but please double check the result."
                    )
                ordered_audio_media = sorted_media

            # Extract only the needed audio media files
            for media in ordered_audio_media:
                ppt_zip.extract(media, temp_dir)

        audio_files = [temp_dir / m for m in ordered_audio_media]

        log(f"Found {len(audio_files)} audio file(s), in slide order:")
        for audio in audio_files:
            log(f"  - {audio.name}")

        # Create an FFmpeg concat file
        concat_file = temp_dir / "concat.txt"
        with open(concat_file, "w", encoding="utf-8") as f:
            for audio in audio_files:
                path = str(audio.resolve()).replace("\\", "/")
                f.write(f"file '{path}'\n")

        # Combine and convert to MP3
        # Check ffmpeg availability first
        ffmpeg_path = shutil.which("ffmpeg")
        if not ffmpeg_path:
            raise RuntimeError(
                "FFmpeg not found in PATH. Install FFmpeg (https://ffmpeg.org/download.html) "
                "and ensure 'ffmpeg' is available on PATH. On Windows: winget install ffmpeg, "
                "on Linux: apt install ffmpeg, on macOS: brew install ffmpeg."
            )

        # Also check for libmp3lame encoder
        probe_result = subprocess.run(
            [ffmpeg_path, "-encoders"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if "libmp3lame" not in probe_result.stdout:
            raise RuntimeError(
                "FFmpeg found but 'libmp3lame' encoder is missing. "
                "Install a full FFmpeg build with MP3 encoding support."
            )

        command = [
            ffmpeg_path,
            "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(concat_file),
            "-vn",
            "-acodec", "libmp3lame",
            "-b:a", "192k",
            str(output_file)
        ]

        log("Combining audio...")

        run_kwargs = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "text": True,
        }
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(command, **run_kwargs)

        if result.returncode != 0:
            log(result.stderr)
            raise RuntimeError("FFmpeg failed to create the MP3.")

        log(f"Done! Output: {output_file.resolve()}")

    return output_file


if __name__ == "__main__":
    extract_audio_to_mp3(
        "eghdamat/1.pptx",
        "eghdamat/1.mp3"
    )
