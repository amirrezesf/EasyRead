

import os
import shutil
import subprocess
import sys
from pathlib import Path


def vector_to_jpg(vector_path, output_path=None, dpi=300):
    vector_path = Path(vector_path)

    if vector_path.suffix.lower() not in {".emf", ".wmf"}:
        raise ValueError("Input file must be an EMF or WMF file.")

    if output_path is None:
        output_path = vector_path.with_suffix(".jpg")
    else:
        output_path = Path(output_path)

    # Convert the vector image to PDF using LibreOffice
    pdf_path = vector_path.with_suffix(".pdf")

    def find_soffice_executable():
        # Prefer executables on PATH first
        for name in ("libreoffice", "soffice"):
            p = shutil.which(name)
            if p:
                return p

        # Common Windows install locations
        if sys.platform.startswith("win"):
            possible_roots = [os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)"), r"C:\Program Files", r"C:\Program Files (x86)"]
            for root in filter(None, possible_roots):
                candidate = Path(root) / "LibreOffice" / "program" / "soffice.exe"
                if candidate.exists():
                    return str(candidate)

        return None

    soffice = find_soffice_executable()
    if not soffice:
        raise FileNotFoundError(
            "LibreOffice (soffice) not found. Install LibreOffice and add it to PATH,\n"
            "or set the PROGRAMFILES environment variable so the script can locate it.\n"
            "Windows users: check 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'.\n"
            "Alternatively, install ImageMagick or another tool that can convert EMF/WMF to PDF/JPG."
        )

    subprocess.run([
        str(soffice),
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(vector_path.parent),
        str(vector_path)
    ], check=True)

    # Convert PDF page to JPG using PyMuPDF
    import fitz

    doc = fitz.open(pdf_path)
    page = doc[0]

    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)

    pix.save(str(output_path))

    doc.close()

    # Remove temporary PDF
    pdf_path.unlink(missing_ok=True)