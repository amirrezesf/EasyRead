"""
pptx_to_html.py

Extract a .pptx deck into a single HTML file that preserves formatting cues
(bold / underline / font color / highlight) as semantic emphasis markup, and
extracts embedded images to a folder with <img> references placed at their
approximate position in reading order.

Usage:
    python pptx_to_html.py input.pptx output.html [--images-dir images]

Design notes / limitations (read before wiring this into your app):

1. "Reading order" on a slide is a heuristic. PPTX shapes have no native
   document order — only a z-order (creation order) and (x, y) coordinates.
   This script sorts shapes by (top, left), which matches simple lecture
   slides (title, then stacked bullets/images) but will misorder anything
   with side-by-side columns or free-form layouts. Spot check a few slides
   before trusting it on your whole deck.

2. "Important" formatting is inferred from bold / underline / font color /
   highlight. This is only as reliable as the professor's own habits — if
   color/underline was used decoratively (or inconsistently) rather than to
   mark importance, this will produce false positives/negatives. Skim the
   HTML output for a slide or two to sanity check before trusting it across
   the whole deck.

3. Images are extracted as files only — this script does NOT caption or
   interpret them. The <img> tag's alt text is a generic placeholder. If you
   want the downstream LLM to actually understand a chart/diagram, you need
   a separate vision-model pass per image that writes a real caption into
   the alt text or a following <p class="img-caption"> tag before you hand
   the HTML to a text-only model. A text-only LLM reading this HTML sees the
   filename, nothing else.

4. Slide notes (speaker notes) are not extracted. Add `slide.notes_slide`
   handling if you need them.
"""

import argparse
import os
import sys
from html import escape

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu


# ---------- formatting detection ----------

def _run_highlight_rgb(run):
    """python-pptx has no public API for <a:highlight> (text highlight color).
    Reach into the underlying XML for it."""
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
    """Bold, underline, an explicit font color, or a highlight all count as
    an 'this looks important' signal."""
    bold = bool(run.font.bold)
    underline = bool(run.font.underline)
    color = _run_font_rgb(run)
    highlight = _run_highlight_rgb(run)
    # Ignore plain black/white font colors — decks often set those explicitly
    # without meaning to flag anything.
    color_is_signal = color is not None and color.upper() not in ("000000", "FFFFFF")
    return bold or underline or color_is_signal or bool(highlight), {
        "bold": bold,
        "underline": underline,
        "color": color if color_is_signal else None,
        "highlight": highlight,
    }


def _run_to_html(run):
    text = escape(run.text)
    if not text:
        return ""
    _, flags = _run_is_emphasized(run)
    if flags["highlight"]:
        text = f'<mark style="background-color:#{flags["highlight"]}">{text}</mark>'
    elif flags["color"]:
        text = f'<span style="color:#{flags["color"]}">{text}</span>'
    if flags["bold"]:
        text = f"<strong>{text}</strong>"
    if flags["underline"]:
        text = f"<u>{text}</u>"
    return text


def _paragraph_to_html(paragraph):
    runs_html = "".join(_run_to_html(r) for r in paragraph.runs)
    if not runs_html.strip():
        return None
    level = paragraph.level or 0
    return level, runs_html


# ---------- shape handlers ----------

def _text_frame_to_html(text_frame):
    parts = []
    list_stack = []  # track open <ul> nesting by level

    def close_to(level):
        while list_stack and list_stack[-1] >= level:
            parts.append("</ul>")
            list_stack.pop()

    any_list = False
    for p in text_frame.paragraphs:
        result = _paragraph_to_html(p)
        if result is None:
            continue
        level, html = result
        close_to(level) if list_stack and list_stack[-1] > level else None
        if not list_stack or list_stack[-1] < level:
            parts.append("<ul>")
            list_stack.append(level)
            any_list = True
        elif list_stack[-1] > level:
            close_to(level)
            if not list_stack:
                parts.append("<ul>")
                list_stack.append(level)
        parts.append(f"<li>{html}</li>")
    while list_stack:
        parts.append("</ul>")
        list_stack.pop()
    return "".join(parts) if any_list else ""


def _table_to_html(table):
    rows_html = []
    for row in table.rows:
        cells_html = []
        for cell in row.cells:
            cell_html = _text_frame_to_html(cell.text_frame).replace("<ul>", "").replace("</ul>", "").replace("<li>", "").replace("</li>", " ")
            cells_html.append(f"<td>{cell_html.strip()}</td>")
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
    return f"<table border='1' cellspacing='0' cellpadding='4'>{''.join(rows_html)}</table>"


def _picture_to_html(shape, slide_idx, img_counter, images_dir, html_out_dir):
    image = shape.image
    ext = image.ext
    filename = f"slide{slide_idx}_img{img_counter}.{ext}"
    filepath = os.path.join(images_dir, filename)
    with open(filepath, "wb") as f:
        f.write(image.blob)
    # path written into the HTML, relative to the HTML file's own directory
    rel_path = os.path.relpath(filepath, html_out_dir)
    return (
        f'<figure class="slide-image">'
        f'<img src="{rel_path}" alt="[TODO: caption slide {slide_idx} image {img_counter} — '
        f'not auto-generated, see script docstring]">'
        f'</figure>'
    )


# ---------- slide / deck assembly ----------

def _shape_sort_key(shape):
    top = shape.top if shape.top is not None else 0
    left = shape.left if shape.left is not None else 0
    return (top, left)


def _slide_to_html(slide, slide_idx, images_dir, html_out_dir):
    pieces = [f'<section class="slide" data-slide="{slide_idx}">', f"<h2>Slide {slide_idx}</h2>"]
    img_counter = 0

    shapes = sorted(slide.shapes, key=_shape_sort_key)

    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            img_counter += 1
            pieces.append(_picture_to_html(shape, slide_idx, img_counter, images_dir, html_out_dir))
            continue

        if getattr(shape, "has_table", False) and shape.has_table:
            pieces.append(_table_to_html(shape.table))
            continue

        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            html = _text_frame_to_html(shape.text_frame)
            if html:
                pieces.append(html)
            continue

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # Recurse into grouped shapes (common for diagrams built from boxes/lines)
            for sub in sorted(shape.shapes, key=_shape_sort_key):
                if sub.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    img_counter += 1
                    pieces.append(_picture_to_html(sub, slide_idx, img_counter, images_dir, html_out_dir))
                elif getattr(sub, "has_text_frame", False) and sub.has_text_frame:
                    html = _text_frame_to_html(sub.text_frame)
                    if html:
                        pieces.append(html)

    pieces.append("</section>")
    return "\n".join(pieces)


def convert(pptx_path, html_path, images_dir=None):
    prs = Presentation(pptx_path)
    html_out_dir = os.path.dirname(os.path.abspath(html_path)) or "."
    if images_dir is None:
        images_dir = os.path.join(html_out_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    slides_html = []
    for i, slide in enumerate(prs.slides, start=1):
        slides_html.append(_slide_to_html(slide, i, images_dir, html_out_dir))

    doc = f"""<!DOCTYPE html>
<html lang="fa">
<head>
<meta charset="UTF-8">
<title>{escape(os.path.basename(pptx_path))}</title>
<style>
  body {{ font-family: sans-serif; direction: rtl; }}
  mark {{ padding: 0 2px; }}
  .slide-image img {{ max-width: 500px; }}
  table {{ border-collapse: collapse; margin: 8px 0; }}
</style>
</head>
<body>
{os.linesep.join(slides_html)}
</body>
</html>
"""
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(doc)

    return html_path, images_dir


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx_path")
    ap.add_argument("html_path")
    ap.add_argument("--images-dir", default=None, help="Defaults to <html_dir>/images")
    args = ap.parse_args()

    html_path, images_dir = convert(args.pptx_path, args.html_path, args.images_dir)
    print(f"Wrote {html_path}")
    print(f"Images in {images_dir}")


if __name__ == "__main__":
    main()
