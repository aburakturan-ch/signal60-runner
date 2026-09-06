from __future__ import annotations

import os
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import List

import pymupdf as fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


@dataclass
class PdfConversionResult:
    source: Path
    output: Path
    pages: int = 0
    paragraphs: int = 0
    images: int = 0
    skipped: bool = False
    message: str = ""
    warning: str = ""


def output_path_for(source: Path) -> Path:
    return source.with_name(f"{source.stem}_Aktarma.docx")


def collect_pdf_from_folder(folder: os.PathLike | str, recursive: bool = False) -> List[Path]:
    folder = Path(folder).expanduser().resolve()
    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted([p for p in folder.glob(pattern) if p.is_file()], key=lambda p: str(p).lower())


def _font_name(raw: str) -> str:
    name = (raw or "Times New Roman").split("+")[-1]
    name = re.sub(r"[-, ](?:BoldItalic|BoldOblique|Bold|Italic|Oblique|Regular|Medium|Light|SemiBold|Black|Book)$", "", name, flags=re.I)
    aliases = {
        "TimesNewRomanPSMT": "Times New Roman",
        "TimesNewRoman": "Times New Roman",
        "ArialMT": "Arial",
        "Helvetica": "Arial",
        "Calibri": "Calibri",
        "Cambria": "Cambria",
        "Georgia": "Georgia",
        "Garamond": "Garamond",
    }
    compact = re.sub(r"[^A-Za-z0-9]", "", name)
    return aliases.get(compact, name.strip() or "Times New Roman")


def _is_bold(span: dict) -> bool:
    f = (span.get("font") or "").lower()
    return bool(span.get("flags", 0) & fitz.TEXT_FONT_BOLD) or "bold" in f


def _is_italic(span: dict) -> bool:
    f = (span.get("font") or "").lower()
    return bool(span.get("flags", 0) & fitz.TEXT_FONT_ITALIC) or "italic" in f or "oblique" in f


def _rgb(value: int) -> RGBColor:
    return RGBColor((value >> 16) & 255, (value >> 8) & 255, value & 255)


def _line_text(line: dict) -> str:
    return "".join(s.get("text", "") for s in line.get("spans", [])).strip()


def _alignment(line: dict, page_width: float):
    bbox = line.get("bbox", (0, 0, 0, 0))
    left, right = bbox[0], bbox[2]
    width = max(1.0, right - left)
    left_gap = left
    right_gap = page_width - right
    if width < page_width * 0.72 and abs(left_gap - right_gap) < page_width * 0.08:
        return WD_ALIGN_PARAGRAPH.CENTER
    if left_gap > page_width * 0.45 and width < page_width * 0.5:
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


def _add_line(doc: Document, line: dict, page_width: float) -> bool:
    if not _line_text(line):
        return False
    p = doc.add_paragraph()
    p.alignment = _alignment(line, page_width)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    for span in line.get("spans", []):
        text = span.get("text", "")
        if not text:
            continue
        run = p.add_run(text)
        run.bold = _is_bold(span)
        run.italic = _is_italic(span)
        run.font.name = _font_name(span.get("font", ""))
        size = float(span.get("size", 11) or 11)
        if 4 <= size <= 72:
            run.font.size = Pt(size)
        run.font.color.rgb = _rgb(int(span.get("color", 0) or 0))
    return True


def _insert_images(doc: Document, page, page_width_inches: float) -> int:
    added = 0
    seen = set()
    for info in page.get_images(full=True):
        xref = info[0]
        if xref in seen:
            continue
        seen.add(xref)
        try:
            pix = fitz.Pixmap(page.parent, xref)
            if pix.width < 40 or pix.height < 40:
                continue
            if pix.alpha or pix.n > 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            data = pix.tobytes("png")
            width = min(page_width_inches * 0.85, max(1.2, pix.width / 110.0))
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(BytesIO(data), width=Inches(width))
            added += 1
        except Exception:
            continue
    return added


def convert_pdf_to_word(source: os.PathLike | str, include_images: bool = False) -> PdfConversionResult:
    source = Path(source).expanduser().resolve()
    output = output_path_for(source)
    if not source.exists() or source.suffix.lower() != ".pdf":
        return PdfConversionResult(source, output, skipped=True, message="PDF bulunamadı veya desteklenmiyor.")

    pdf = None
    try:
        pdf = fitz.open(source)
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

        paragraphs = 0
        images = 0
        text_chars = 0
        for page_index, page in enumerate(pdf):
            data = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            lines = []
            for block in data.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    txt = _line_text(line)
                    if txt:
                        lines.append((line.get("bbox", (0, 0, 0, 0))[1], line.get("bbox", (0, 0, 0, 0))[0], line))
                        text_chars += len(txt)
            lines.sort(key=lambda item: (round(item[0], 1), item[1]))
            for _, __, line in lines:
                if _add_line(doc, line, page.rect.width):
                    paragraphs += 1
            if include_images:
                usable_width = (section.page_width - section.left_margin - section.right_margin) / 914400
                images += _insert_images(doc, page, usable_width)
            if page_index < len(pdf) - 1:
                doc.add_page_break()

        warning = ""
        if text_chars < max(40, len(pdf) * 15):
            warning = "PDF'de yeterli metin katmanı bulunamadı. Bu dosya taranmış görüntü olabilir; OCR uygulanmadı."
        doc.save(output)
        return PdfConversionResult(source, output, pages=len(pdf), paragraphs=paragraphs, images=images, message="Tamamlandı.", warning=warning)
    except Exception as exc:
        return PdfConversionResult(source, output, skipped=True, message=f"Hata: {exc}")
    finally:
        if pdf is not None:
            pdf.close()
