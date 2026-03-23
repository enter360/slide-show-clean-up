#!/usr/bin/env python3
"""
Remove blank slides from .pptx and blank pages from .pdf files.

Usage:
    python remove_blank_slides.py

Drop files into the 'input' folder and run the script.
Cleaned files will appear in the 'output' folder.

Dependencies:
    pip install python-pptx pypdf pdfplumber Pillow
"""

import sys
from pathlib import Path


INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")


# ---------------------------------------------------------------------------
# PPTX helpers
# ---------------------------------------------------------------------------

def slide_is_blank(slide, text_threshold: int = 3) -> bool:
    from pptx.dml.color import RGBColor

    for shape in slide.shapes:
        if shape.shape_type == 13:
            return False
        if hasattr(shape, "image"):
            return False
        if shape.has_text_frame:
            if len(shape.text_frame.text.strip()) > text_threshold:
                return False
        try:
            fill = shape.fill
            if fill.type is not None:
                if fill.type.name == "SOLID":
                    try:
                        if fill.fore_color.rgb != RGBColor(0xFF, 0xFF, 0xFF):
                            return False
                    except Exception:
                        return False
                else:
                    return False
        except Exception:
            pass
        if shape.shape_type == 6 and len(shape.shapes) > 0:
            return False

    return True


def remove_blank_slides_pptx(input_path: Path, output_path: Path) -> dict:
    from pptx import Presentation

    prs = Presentation(str(input_path))
    xml_slides = prs.slides._sldIdLst
    total = len(prs.slides)
    blank_indices = [i for i, s in enumerate(prs.slides) if slide_is_blank(s)]

    for i in sorted(blank_indices, reverse=True):
        rId = prs.slides._sldIdLst[i].get("r:id")
        prs.part.drop_rel(rId)
        del xml_slides[i]

    prs.save(str(output_path))
    return {"total": total, "removed": len(blank_indices)}


# ---------------------------------------------------------------------------
# PDF helpers
# ---------------------------------------------------------------------------

def pdf_page_is_blank(page, text_threshold: int = 3) -> bool:
    if len((page.extract_text() or "").strip()) > text_threshold:
        return False
    if page.images or page.lines or page.curves or page.rects:
        return False
    return True


def remove_blank_pages_pdf(input_path: Path, output_path: Path) -> dict:
    import pdfplumber
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(input_path))
    total = len(reader.pages)

    with pdfplumber.open(str(input_path)) as pdf:
        blank_indices = {i for i, p in enumerate(pdf.pages) if pdf_page_is_blank(p)}

    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i not in blank_indices:
            writer.add_page(page)
    if reader.metadata:
        writer.add_metadata(reader.metadata)

    with open(str(output_path), "wb") as f:
        writer.write(f)

    return {"total": total, "removed": len(blank_indices)}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    files = [f for f in INPUT_DIR.iterdir() if f.suffix.lower() in {".pptx", ".pdf"}]

    if not files:
        print(f"No .pptx or .pdf files found in '{INPUT_DIR}/'.")
        print("Add files to the input folder and run again.")
        return

    try:
        for input_path in sorted(files):
            output_path = OUTPUT_DIR / input_path.name
            suffix = input_path.suffix.lower()

            print(f"Processing: {input_path.name}")

            if suffix == ".pptx":
                result = remove_blank_slides_pptx(input_path, output_path)
                label = "slides"
            else:
                result = remove_blank_pages_pdf(input_path, output_path)
                label = "pages"

            print(f"  {result['total']} {label} total, {result['removed']} blank removed")
            print(f"  Saved to: {output_path}\n")

    except ImportError as e:
        sys.exit(
            f"Missing dependency: {e}\n"
            "Install with:  pip install python-pptx pypdf pdfplumber Pillow"
        )


if __name__ == "__main__":
    main()
