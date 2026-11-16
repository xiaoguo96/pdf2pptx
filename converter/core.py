# -*- coding: utf-8 -*-
"""PDF -> editable PPTX core pipeline.

Strategy
--------
Image-only PDFs (scanned / exported-as-picture slides) are converted by:

1. Rendering each page to a high-resolution raster.
2. Running OCR (RapidOCR, PaddleOCR models on ONNX runtime) to obtain
   per-line text, bounding boxes and confidence.
3. For every text line: estimating font size, color, boldness and
   alignment from the raster pixels inside its bounding box.
4. Inpainting the text pixels out of the page raster so the background
   stays visually intact but contains no baked-in text.
5. Building a PPTX slide whose size equals the PDF page size:
   - the cleaned raster is inserted as one full-bleed Picture object
     (selectable / movable / resizable);
   - every text line becomes a native editable TextBox with matched
     position, size, color, bold and font.

Known limitation: decorative vector artwork, charts and tables that are
part of the page raster remain inside the background picture (they are
not natively reconstructable from a raster source). See README.
"""

import io
import json
import os
import tempfile
from dataclasses import dataclass, field, asdict

import cv2
import numpy as np
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn

EMU_PER_PT = 12700
CJK_FONT = "Microsoft YaHei"
LATIN_FONT = "Microsoft YaHei"


# --------------------------------------------------------------------------
# OCR wrapper
# --------------------------------------------------------------------------
class OcrEngine:
    def __init__(self):
        from rapidocr_onnxruntime import RapidOCR
        self._ocr = RapidOCR()

    def lines(self, img_bgr):
        """Return list of (quad, text, score); quad = 4 pts in pixel space."""
        res, _ = self._ocr(img_bgr)
        out = []
        for r in res or []:
            quad, text, score = r[0], r[1], float(r[2])
            xs = [p[0] for p in quad]
            ys = [p[1] for p in quad]
            out.append({
                "quad": [[round(float(p[0]), 2), round(float(p[1]), 2)] for p in quad],
                "bbox": [min(xs), min(ys), max(xs), max(ys)],
                "text": text,
                "score": round(score, 3),
            })
        return out


# --------------------------------------------------------------------------
# Per-line typography analysis
# --------------------------------------------------------------------------
@dataclass
class TextLine:
    text: str
    score: float
    x: float          # px, in render space
    y: float
    w: float
    h: float
    size_pt: float
    color: tuple      # RGB 0-255
    bold: bool
    mask: np.ndarray = field(repr=False, default=None)  # uint8 0/255, crop space


def _binarize(crop_bgr):
    """Return (mask, polarity) where mask>0 selects TEXT pixels."""
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    ret, m = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # text pixels are the minority component inside a tight text crop
    if float((m > 0).mean()) > 0.5:
        m = cv2.bitwise_not(m)
    return m


def _text_color(crop_bgr, mask):
    """Median color of the core (most saturated vs background) text pixels."""
    ys, xs = np.where(mask > 0)
    if len(xs) == 0:
        return (255, 255, 255)
    px = crop_bgr[ys, xs].astype(np.int32)
    bg = np.median(px, axis=0)
    # distance from overall median: core glyph pixels deviate most
    d = np.abs(px - bg).sum(axis=1)
    thr = np.percentile(d, 55)
    core = px[d >= thr]
    if len(core) == 0:
        core = px
    b, g, r = [int(v) for v in np.median(core, axis=0)]
    return (r, g, b)


def _is_bold(mask, h):
    """Estimate stroke width / height ratio via distance transform."""
    if h <= 2 or mask is None:
        return False
    dist = cv2.distanceTransform(mask, cv2.DIST_L2, 3)
    vals = dist[mask > 0]
    if len(vals) == 0:
        return False
    stroke = float(np.percentile(vals, 75)) * 2.0
    return (stroke / h) > 0.13


def analyze_lines(img_bgr, ocr_lines, dpi):
    """Attach typography attributes to OCR lines."""
    px_to_pt = 72.0 / dpi
    out = []
    for ln in ocr_lines:
        x0, y0, x1, y1 = ln["bbox"]
        x0 = max(0, int(np.floor(x0)))
        y0 = max(0, int(np.floor(y0)))
        x1 = min(img_bgr.shape[1], int(np.ceil(x1)))
        y1 = min(img_bgr.shape[0], int(np.ceil(y1)))
        if x1 - x0 < 2 or y1 - y0 < 2:
            continue
        crop = img_bgr[y0:y1, x0:x1]
        mask = _binarize(crop)
        h_px = y1 - y0
        # tight OCR box ≈ cap/glyph height ≈ 0.86 em for CJK
        size_pt = max(7.0, round(h_px * px_to_pt / 0.86, 1))
        out.append(TextLine(
            text=ln["text"],
            score=ln["score"],
            x=x0, y=y0, w=x1 - x0, h=y1 - y0,
            size_pt=size_pt,
            color=_text_color(crop, mask),
            bold=_is_bold(mask, h_px),
            mask=mask,
        ))
    return out


# --------------------------------------------------------------------------
# Background restoration (text inpainting)
# --------------------------------------------------------------------------
def inpaint_text(img_bgr, lines, radius=4):
    """Remove text pixels from the page raster, keep the background."""
    h, w = img_bgr.shape[:2]
    full = np.zeros((h, w), np.uint8)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    for ln in lines:
        if ln.mask is None:
            continue
        mh, mw = ln.mask.shape
        py1 = min(h, ln.y + mh)
        px1 = min(w, ln.x + mw)
        sub = ln.mask[0:py1 - ln.y, 0:px1 - ln.x]
        region = full[ln.y:py1, ln.x:px1]
        np.maximum(region, sub, out=region)
    # expand mask edges to cover anti-aliased halos
    full = cv2.dilate(full, kernel, iterations=2)
    fixed = cv2.inpaint(img_bgr, full, radius, cv2.INPAINT_TELEA)
    return fixed, full


# --------------------------------------------------------------------------
# PPTX building
# --------------------------------------------------------------------------
def _set_run_font(run, size_pt, color, bold, text):
    f = run.font
    f.size = Pt(size_pt)
    f.bold = bool(bold)
    f.name = LATIN_FONT
    f.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    # ensure east-asian typeface is set so CJK renders with the same font
    for tag in ("a:ea",):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", CJK_FONT)


def build_pptx(pages, out_path, dpi):
    """pages: list of dicts {bg_bytes(jpeg), lines: [TextLine], w_pt, h_pt}."""
    prs = Presentation()
    manifest = []
    first = pages[0]
    prs.slide_width = Emu(int(round(first["w_pt"] * EMU_PER_PT)))
    prs.slide_height = Emu(int(round(first["h_pt"] * EMU_PER_PT)))
    blank = prs.slide_layouts[6]
    emu_per_px = (72.0 / dpi) * EMU_PER_PT

    for idx, pg in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        bg = io.BytesIO(pg["bg_bytes"])
        pic = slide.shapes.add_picture(bg, 0, 0,
                                       width=prs.slide_width,
                                       height=prs.slide_height)
        pic.name = f"Page{idx + 1} Background"
        entry = {"page": idx + 1, "picture": pic.name, "texts": []}
        for ln in pg["lines"]:
            left = Emu(int(round(ln.x * emu_per_px)))
            top = Emu(int(round(ln.y * emu_per_px)))
            width = Emu(int(round((ln.w + 6) * emu_per_px)))
            height = Emu(int(round((ln.h + 4) * emu_per_px)))
            box = slide.shapes.add_textbox(left, top, width, height)
            box.name = f"Text: {ln.text[:12]}"
            tf = box.text_frame
            tf.word_wrap = False
            tf.margin_left = 0
            tf.margin_right = 0
            tf.margin_top = 0
            tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]
            run = para.add_run()
            run.text = ln.text
            _set_run_font(run, ln.size_pt, ln.color, ln.bold, ln.text)
            entry["texts"].append({
                "text": ln.text, "score": ln.score,
                "x_pt": round(ln.x * 72 / dpi, 2), "y_pt": round(ln.y * 72 / dpi, 2),
                "w_pt": round(ln.w * 72 / dpi, 2), "h_pt": round(ln.h * 72 / dpi, 2),
                "size_pt": ln.size_pt, "color": list(ln.color), "bold": ln.bold,
            })
        manifest.append(entry)

    prs.save(out_path)
    return manifest


# --------------------------------------------------------------------------
# Main entry
# --------------------------------------------------------------------------
def convert_pdf_to_pptx(pdf_path, pptx_path, dpi=150, manifest_path=None,
                        workdir=None, keep_workdir=False):
    doc = pymupdf.open(pdf_path)
    if doc.page_count == 0:
        raise ValueError("PDF has no pages")
    ocr = OcrEngine()
    pages = []
    all_manifest = []
    for idx, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        lines = analyze_lines(bgr, ocr.lines(bgr), dpi)
        fixed, _mask = inpaint_text(bgr, lines)
        ok, buf = cv2.imencode(".jpg", fixed,
                               [int(cv2.IMWRITE_JPEG_QUALITY), 88])
        pages.append({
            "bg_bytes": buf.tobytes(),
            "lines": lines,
            "w_pt": page.rect.width,
            "h_pt": page.rect.height,
        })
    out = build_pptx(pages, pptx_path, dpi)
    doc.close()
    if manifest_path:
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
    return out
