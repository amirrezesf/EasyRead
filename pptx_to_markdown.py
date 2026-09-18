"""
Extract a .pptx deck into a single Markdown file that preserves formatting
cues (bold / underline / font color / highlight) and extracts embedded images
with Markdown image references in reading order.

Usage:
    python pptx_to_markdown.py input.pptx output.md [--images-dir images]
"""

import argparse
import os
import re

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from convert_emf_and_wmf import vector_to_jpg


def _run_highlight_rgb(run):
    """python-pptx has no public API for <a:highlight> (text highlight color)."""
    r = run._r
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    hl = r.find("a:rPr/a:highlight/a:srgbClr", ns)
    if hl is not None:
        return hl.get("val")
    return None


def _run_font_rgb(run):
    try:
        color = run.font.color
        if color is not None and color.type is not None and color.rgb is not None:
            return str(color.rgb)
    except Exception:
        pass
    return None


def _run_is_emphasized(run):
    bold = bool(run.font.bold)
    underline = bool(run.font.underline)
    color = _run_font_rgb(run)
    highlight = _run_highlight_rgb(run)
    color_is_signal = color is not None and color.upper() not in ("000000", "FFFFFF")
    return bold or underline or color_is_signal or bool(highlight), {
        "bold": bold,
        "underline": underline,
        "color": color if color_is_signal else None,
        "highlight": highlight,
    }


def _escape_md(text):
    return (
        text.replace("\\", "\\\\")
        .replace("*", "\\*")
        .replace("_", "\\_")
        .replace("`", "\\`")
        .replace("[", "\\[")
        .replace("]", "\\]")
    )


def _run_to_markdown(run):
    text = _escape_md(run.text)
    if not text:
        return ""
    _, flags = _run_is_emphasized(run)
    if flags["highlight"]:
        text = f"=={text}=="
    elif flags["color"]:
        text = f'<span style="color:#{flags["color"]}">{text}</span>'
    if flags["bold"]:
        text = f"**{text}**"
    if flags["underline"]:
        text = f"<u>{text}</u>"
    return text


def _paragraph_to_markdown(paragraph):
    runs_md = "".join(_run_to_markdown(r) for r in paragraph.runs)
    if not runs_md.strip():
        return None
    level = paragraph.level or 0
    return level, runs_md


def _text_frame_to_markdown(text_frame):
    lines = []
    for p in text_frame.paragraphs:
        result = _paragraph_to_markdown(p)
        if result is None:
            continue
        level, md = result
        indent = "  " * level
        lines.append(f"{indent}- {md}")
    return "\n".join(lines)


def _cell_plain(cell):
    parts = []
    for p in cell.text_frame.paragraphs:
        result = _paragraph_to_markdown(p)
        if result is not None:
            parts.append(result[1].replace("|", "\\|"))
    return " ".join(parts).strip()


def _table_to_markdown(table):
    rows = []
    for row in table.rows:
        cells = [_cell_plain(cell) for cell in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    for row in rows:
        while len(row) < col_count:
            row.append("")
    header = rows[0]
    body = rows[1:] if len(rows) > 1 else []
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    if not body:
        lines.append("| " + " | ".join("" for _ in header) + " |")
    else:
        for row in body:
            lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _picture_to_markdown(shape, slide_idx, img_counter, images_dir, md_out_dir):
    image = shape.image
    ext = image.ext.lower()
    filename = f"slide{slide_idx}_img{img_counter}.{ext}"
    filepath = os.path.join(images_dir, filename)
    with open(filepath, "wb") as f:
        f.write(image.blob)
    if ext in {"emf", "wmf"}:
        converted_path = os.path.splitext(filepath)[0] + ".jpg"
        vector_to_jpg(filepath, converted_path)
        os.remove(filepath)
        filepath = converted_path
    rel_path = os.path.relpath(filepath, md_out_dir).replace("\\", "/")
    alt = f"Slide {slide_idx} image {img_counter}"
    return f"![{alt}]({rel_path})"


def _shape_sort_key(shape):
    top = shape.top if shape.top is not None else 0
    left = shape.left if shape.left is not None else 0
    return (top, left)


def _collect_shape_markdown(shape, slide_idx, img_counter, images_dir, md_out_dir, pieces):
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        img_counter += 1
        pieces.append(
            _picture_to_markdown(shape, slide_idx, img_counter, images_dir, md_out_dir)
        )
        return img_counter

    if getattr(shape, "has_table", False) and shape.has_table:
        md = _table_to_markdown(shape.table)
        if md:
            pieces.append(md)
        return img_counter

    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        md = _text_frame_to_markdown(shape.text_frame)
        if md:
            pieces.append(md)
        return img_counter

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for sub in sorted(shape.shapes, key=_shape_sort_key):
            img_counter = _collect_shape_markdown(
                sub, slide_idx, img_counter, images_dir, md_out_dir, pieces
            )
        return img_counter

    return img_counter


def _slide_to_markdown(slide, slide_idx, images_dir, md_out_dir):
    pieces = [f"## Slide {slide_idx}"]
    img_counter = 0
    for shape in sorted(slide.shapes, key=_shape_sort_key):
        img_counter = _collect_shape_markdown(
            shape, slide_idx, img_counter, images_dir, md_out_dir, pieces
        )
    return "\n\n".join(pieces)


def convert(pptx_path, md_path, images_dir=None, log=print):
    prs = Presentation(pptx_path)
    md_out_dir = os.path.dirname(os.path.abspath(md_path)) or "."
    os.makedirs(md_out_dir, exist_ok=True)
    if images_dir is None:
        images_dir = os.path.join(md_out_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    slides_md = []
    for i, slide in enumerate(prs.slides, start=1):
        slides_md.append(_slide_to_markdown(slide, i, images_dir, md_out_dir))
        log(f"Converted slide {i}/{len(prs.slides)}")

    title = os.path.splitext(os.path.basename(pptx_path))[0]
    doc = f"# {title}\n\n" + "\n\n---\n\n".join(slides_md) + "\n"
    # Collapse excessive blank lines
    doc = re.sub(r"\n{3,}", "\n\n", doc)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(doc)

    return md_path, images_dir


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx_path")
    ap.add_argument("md_path")
    ap.add_argument("--images-dir", default=None, help="Defaults to <md_dir>/images")
    args = ap.parse_args()

    md_path, images_dir = convert(args.pptx_path, args.md_path, args.images_dir)
    print(f"Wrote {md_path}")
    print(f"Images in {images_dir}")


if __name__ == "__main__":
    main()
