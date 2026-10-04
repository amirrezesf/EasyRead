"""Convert EMF/WMF vector images to JPG using available system tools.

Supports LibreOffice (soffice), ImageMagick (magick/convert), and unoconv
as fallback converters. Best results with LibreOffice.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


def _find_executable(
    names: list[str], windows_paths: Optional[list[Path]] = None
) -> Optional[str]:
    """Find an executable in PATH or common Windows install locations."""
    for name in names:
        p = shutil.which(name)
        if p:
            return p
    if sys.platform.startswith("win") and windows_paths:
        for path in windows_paths:
            candidate = Path(path)
            if candidate.exists():
                return str(candidate)
    return None


def _find_soffice() -> Optional[str]:
    """Find LibreOffice/soffice executable."""
    return _find_executable(
        ["libreoffice", "soffice"],
        windows_paths=[
            Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice" / "program" / "soffice.exe",
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ],
    )


def _find_imagemagick() -> Optional[str]:
    """Find ImageMagick convert/magick executable."""
    return _find_executable(
        ["magick", "convert"],
        windows_paths=[
            Path(os.environ.get("PROGRAMFILES", "")) / "ImageMagick*" / "magick.exe",
            Path(os.environ.get("PROGRAMFILES", "")) / "ImageMagick*" / "convert.exe",
            Path(r"C:\Program Files\ImageMagick*\magick.exe"),
            Path(r"C:\Program Files\ImageMagick*\convert.exe"),
        ],
    )


def _find_unoconv() -> Optional[str]:
    """Find unoconv executable (uses LibreOffice headless)."""
    return _find_executable(["unoconv"])


def vector_to_jpg(vector_path: str | Path, output_path: Optional[str | Path] = None, dpi: int = 300) -> None:
    vector_path = Path(vector_path)

    if vector_path.suffix.lower() not in {".emf", ".wmf"}:
        raise ValueError("Input file must be an EMF or WMF file.")

    if output_path is None:
        output_path = vector_path.with_suffix(".jpg")
    else:
        output_path = Path(output_path)

    # Try LibreOffice first (best for EMF/WMF)
    soffice = _find_soffice()
    if soffice:
        try:
            pdf_path = vector_path.with_suffix(".pdf")
            subprocess.run(
                [
                    str(soffice),
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(vector_path.parent),
                    str(vector_path)
                ],
                check=True,
                capture_output=True,
                text=True,
            )

            # Convert PDF page to JPG using PyMuPDF
            import fitz

            doc = fitz.open(pdf_path)
            page = doc[0]

            zoom = dpi / 72
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)

            pix.save(str(output_path))

            doc.close()
            pdf_path.unlink(missing_ok=True)
            return
        except subprocess.CalledProcessError as e:
            print(f"LibreOffice conversion failed: {e.stderr}", file=sys.stderr)
        except Exception as e:
            print(f"LibreOffice path failed: {e}", file=sys.stderr)

    # Fallback: ImageMagick
    magick = _find_imagemagick()
    if magick:
        try:
            subprocess.run(
                [
                    str(magick),
                    str(vector_path),
                    "-density", str(dpi),
                    "-background", "white",
                    "-alpha", "remove",
                    "-alpha", "off",
                    str(output_path)
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            return
        except subprocess.CalledProcessError as e:
            print(f"ImageMagick conversion failed: {e.stderr}", file=sys.stderr)
        except Exception as e:
            print(f"ImageMagick path failed: {e}", file=sys.stderr)

    # Fallback: unoconv
    unoconv = _find_unoconv()
    if unoconv:
        try:
            pdf_path = vector_path.with_suffix(".pdf")
            subprocess.run(
                [str(unoconv), "-f", "pdf", "-o", str(pdf_path), str(vector_path)],
                check=True,
                capture_output=True,
                text=True,
            )

            import fitz
            doc = fitz.open(pdf_path)
            page = doc[0]
            zoom = dpi / 72
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)

            pix.save(str(output_path))
            doc.close()
            pdf_path.unlink(missing_ok=True)
            return
        except Exception as e:
            print(f"unoconv path failed: {e}", file=sys.stderr)

    raise FileNotFoundError(
        "No suitable converter found for EMF/WMF. Install one of:\n"
        "  - LibreOffice (soffice) - recommended\n"
        "  - ImageMagick (magick/convert)\n"
        "  - unoconv (requires LibreOffice)\n"
        "On Windows: winget install TheDocumentFoundation.LibreOffice\n"
        "On Linux: apt install libreoffice imagemagick\n"
        "On macOS: brew install --cask libreoffice imagemagick"
    )