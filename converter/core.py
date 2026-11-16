# -*- coding: utf-8 -*-
"""PDF -> editable PPTX core pipeline.

Strategy
--------
Image-only PDFs (scanned / exported-as-picture slides) are converted by:

1. Rendering each page to a high-resolution raster.
2. Running OCR (RapidOCR) to obtain per-line text, bounding boxes,
   confidence.
3. Classifying every OCR line as PROTECTED or REBUILD:

   * PROTECTED - left completely untouched in the background raster and
     never re-rendered (page numbers, logos, chart values, badge numbers
     etc. keep their original pixels):
       - header / footer bands (y-ratio rules),
       - pure numeric / percent / thousands / year labels,
       - very short small labels,
       - merged OCR lines are split: digit segments stay protected,
         only the middle text segment is rebuilt,
       - any line whose surrounding ring is strongly colored (sits on
         chart bars / icons / artwork) or non-smooth (photo / gradient).

   * REBUILD - normal body text on smooth backgrounds: erased with a
     row-wise linear-interpolation fill (avoids TELEA smearing /
     shadow artifacts) and replaced by a native editable TextBox.
     Font size is constrained by BOTH glyph height and measured line
     width so re-rendered text never overflows its original box / card.

4. Building a PPTX slide whose size equals the PDF page size:
   - the cleaned raster is inserted as one full-bleed Picture object;
   - every REBUILD line becomes a native editable TextBox.
"""

import io
import json
import os
import re
from dataclasses import dataclass, field, asdict, replace

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

# header / footer protection bands (fraction of page height)
HEADER_FRAC = 0.075
FOOTER_FRAC = 0.92
# ring background checks
RING_GRAD_MAX = 30.0     # 90th-percentile Sobel magnitude -> "smooth"
RING_SAT_FRAC_MAX = 0.30  # fraction of strongly colored ring pixels


# --------------------------------------------------------------------------
# OCR wrapper
# --------------------------------------------------------------------------
class OcrEngine:
    def __init__(self):
        from rapidocr_onnxruntime import RapidOCR
        self._ocr = RapidOCR()

    def lines(self, img_bgr):
        """Return list of dicts with quad/bbox/text/score (pixel space)."""
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
    """Return mask where mask>0 selects TEXT pixels."""
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    ret, m = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # text pixels are the minority component inside a tight text crop
    if float((m > 0).mean()) > 0.5:
        m = cv2.bitwise_not(m)
    return m


def _text_color(crop_bgr, mask):
    """Median color of the core (most deviating) text pixels."""
    ys, xs = np.where(mask > 0)
    if len(xs) == 0:
        return (255, 255, 255)
    px = crop_bgr[ys, xs].astype(np.int32)
    bg = np.median(px, axis=0)
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
        # tight OCR box ~ cap/glyph height ~ 0.86 em for CJK
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
# Line classification: PROTECT vs REBUILD
# --------------------------------------------------------------------------
RE_NUMERIC = re.compile(r"[\d][\d,.:%\-–~— ]{0,11}")
RE_LEAD_DIGITS = re.compile(r"^(\d{2,3})(?=\D)")
RE_TRAIL_DIGITS = re.compile(r"(\d{1,3}(?:\.\d+)?%?)$")


def _is_numericish(t):
    t = t.strip()
    if not t or len(t) > 12 or not any(c.isdigit() for c in t):
        return False
    return RE_NUMERIC.fullmatch(t) is not None


def _char_w(ch):
    """Approximate advance width in em units (YaHei-like metrics)."""
    o = ord(ch)
    if o >= 0x2E80 or ch in "\u00b7\u2014\u2026":
        return 1.0
    if ch.isdigit():
        return 0.55
    if ch.isalpha():
        return 0.72 if ch.isupper() else 0.52
    if ch == " ":
        return 0.30
    return 0.40


def _factor(text):
    return sum(_char_w(c) for c in text) or 1.0


def _ring_stats(img, x0, y0, x1, y1, pad=12, gap=4):
    """Gradient energy + color saturation of the LEFT/RIGHT side bands.

    Only horizontal bands are checked (matching the row-wise fill which
    interpolates from left/right neighbours); a `gap` margin next to the
    bbox is skipped so glyph anti-aliasing halos don't count as edges.
    """
    H, W = img.shape[:2]
    y0c, y1c = max(0, y0), min(H, y1)
    if y1c - y0c < 3:
        return 1e9, 1.0
    bands = []
    la, lb = max(0, x0 - pad), max(0, x0 - gap)
    if lb - la >= 3:
        bands.append(img[y0c:y1c, la:lb])
    ra, rb = min(W, x1 + gap), min(W, x1 + pad)
    if rb - ra >= 3:
        bands.append(img[y0c:y1c, ra:rb])
    if not bands:
        return 1e9, 1.0
    px = np.concatenate([b.reshape(-1, 3) for b in bands], axis=0).astype(np.int16)
    if px.shape[0] < 8:
        return 1e9, 1.0
    mags = []
    for b in bands:
        g = cv2.cvtColor(b, cv2.COLOR_BGR2GRAY)
        gx = cv2.Sobel(g, cv2.CV_32F, 1, 0)
        gy = cv2.Sobel(g, cv2.CV_32F, 0, 1)
        mags.append(np.sqrt(gx * gx + gy * gy).ravel())
    mag = np.concatenate(mags)
    sat = px.max(axis=1) - px.min(axis=1)
    frac_sat = float((sat > 80).mean())
    grad90 = float(np.percentile(mag, 90))
    return grad90, frac_sat


def _split_merged(img, ln):
    """Split '01Title03'-style merged OCR lines.

    Returns (line_or_None, changed). The digit segments keep their
    original pixels (protected); only the middle text becomes a new
    TextLine with a proportional bbox.
    """
    t = ln.text.strip()
    lead = RE_LEAD_DIGITS.match(t)
    core, cx, cw, changed = t, float(ln.x), float(ln.w), False
    if lead:
        pre = lead.group(1)
        core = t[len(pre):]
        f_pre, f_core = _factor(pre), _factor(core)
        dx = cw * f_pre / (f_pre + f_core)
        cx, cw = cx + dx, cw - dx
        changed = True
    trail = RE_TRAIL_DIGITS.search(core)
    if trail and trail.start() >= 3:
        suf = trail.group(1)
        core = core[:trail.start()]
        f_suf, f_core = _factor(suf), _factor(core)
        cw = cw * f_core / (f_core + f_suf)
        changed = True
    if not changed or len(core) < 3 or cw < 10:
        return ln, False
    x0 = max(0, int(np.floor(cx)))
    y0 = int(ln.y)
    x1 = min(img.shape[1], int(np.ceil(cx + cw)))
    y1 = int(ln.y + ln.h)
    if x1 - x0 < 6:
        return ln, False
    crop = img[y0:y1, x0:x1]
    mask = _binarize(crop)
    return replace(ln, text=core, x=x0, w=x1 - x0, mask=mask), True


def _fit_size_pt(ln, dpi):
    """Font size constrained by BOTH glyph height and measured width."""
    avail_pt = ln.w * 72.0 / dpi
    size_w = avail_pt / _factor(ln.text)
    return max(7.0, round(min(ln.size_pt, size_w), 1))


def classify_lines(img, lines, dpi, page_h_px):
    """Return (rebuild_lines, protected_infos)."""
    rebuild, protected = [], []
    for ln in lines:
        t = ln.text.strip()
        if not t:
            protected.append({"text": ln.text, "reason": "empty"})
            continue
        # header / footer bands (logo + page numbers)
        if ln.y + ln.h < HEADER_FRAC * page_h_px:
            protected.append({"text": ln.text, "reason": "header band"})
            continue
        if ln.y > FOOTER_FRAC * page_h_px:
            protected.append({"text": ln.text, "reason": "footer band"})
            continue
        # pure numeric / percent / year labels (chart values, badges, TOC)
        if _is_numericish(t):
            protected.append({"text": ln.text, "reason": "numeric label"})
            continue
        # very short small labels
        if len(t) <= 3 and ln.h <= 0.02 * page_h_px:
            protected.append({"text": ln.text, "reason": "short small label"})
            continue
        # merged "01Title03" -> keep digit segments, rebuild middle text
        ln, changed = _split_merged(img, ln)
        # background checks around (possibly split) bbox
        grad90, frac_sat = _ring_stats(img, int(ln.x), int(ln.y),
                                       int(ln.x + ln.w), int(ln.y + ln.h))
        if frac_sat > RING_SAT_FRAC_MAX:
            protected.append({"text": ln.text, "reason": "on colored graphic"})
            continue
        if grad90 > RING_GRAD_MAX:
            protected.append({"text": ln.text, "reason": "complex background"})
            continue
        ln = replace(ln, size_pt=_fit_size_pt(ln, dpi))
        rebuild.append(ln)
    return rebuild, protected


# --------------------------------------------------------------------------
# Background restoration (smooth text removal, no TELEA shadow)
# --------------------------------------------------------------------------
def _fill_masked_rows(crop, mask):
    """Row-wise linear interpolation between nearest unmasked pixels."""
    h, w = mask.shape
    m = mask > 0
    pos = np.arange(w)
    full_rows = []
    for r in range(h):
        row = m[r]
        if not row.any():
            continue
        if row.all():
            full_rows.append(r)
            continue
        um = ~row
        lx = np.maximum.accumulate(np.where(um, pos, -1))
        rx = np.minimum.accumulate(np.where(um, pos, w)[::-1])[::-1]
        xs = np.flatnonzero(row)
        l_ok = lx[xs] >= 0
        r_ok = rx[xs] < w
        span = (rx[xs] - lx[xs]).astype(np.float64)
        t = np.ones(len(xs))
        both = l_ok & r_ok & (span > 0)
        t[both] = (xs[both] - lx[xs[both]]) / span[both]
        t[l_ok & ~r_ok] = 0.0
        src_l = np.where(l_ok, lx[xs], rx[xs])
        src_r = np.where(r_ok, rx[xs], lx[xs])
        interp = ((1.0 - t)[:, None] * crop[r, src_l].astype(np.float64)
                  + t[:, None] * crop[r, src_r].astype(np.float64))
        crop[r, xs] = np.clip(interp, 0, 255).astype(np.uint8)
    if full_rows:
        clean = [r for r in range(h) if not m[r].all()]
        if clean:
            ca = np.array(clean)
            for r in full_rows:
                above = ca[ca < r]
                below = ca[ca > r]
                if len(above) and len(below):
                    a, b = int(above[-1]), int(below[0])
                    tt = (r - a) / float(b - a)
                    crop[r] = ((1 - tt) * crop[a] + tt * crop[b]).astype(np.uint8)
                elif len(above):
                    crop[r] = crop[int(above[-1])]
                else:
                    crop[r] = crop[int(below[0])]


def erase_lines(img_bgr, lines):
    """Remove REBUILD line pixels via smooth fill. Modifies a copy."""
    out = img_bgr.copy()
    h, w = out.shape[:2]
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    for ln in lines:
        if ln.mask is None:
            continue
        mh, mw = ln.mask.shape
        py1 = min(h, ln.y + mh)
        px1 = min(w, ln.x + mw)
        m = cv2.dilate(ln.mask[0:py1 - ln.y, 0:px1 - ln.x], kernel, iterations=2)
        _fill_masked_rows(out[ln.y:py1, ln.x:px1], m)
    return out


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
        entry = {"page": idx + 1, "picture": pic.name, "texts": [],
                 "protected": pg.get("protected", []),
                 "source": pg.get("source", "ocr(RapidOCR)")}
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
# LLM (vision API) based layout extraction
# --------------------------------------------------------------------------
def _hex_to_rgb(hexstr):
    s = (hexstr or "").strip().lstrip("#")
    if len(s) != 6 or not all(c in "0123456789abcdefABCDEF" for c in s):
        return None
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


def lines_from_llm(img_bgr, elements, dpi):
    """Convert LLM elements to (rebuild_lines, protected_infos).

    kind != "body"  -> PROTECT (original pixels kept, no erase).
    kind == "body"  -> split into lines, rebuild as editable text boxes.
    """
    rebuild, protected = [], []
    H, W = img_bgr.shape[:2]
    px_to_pt = 72.0 / dpi
    for el in elements:
        if el.kind != "body":
            protected.append({
                "text": el.text.replace("\n", " ")[:40],
                "reason": f"llm:{el.kind}",
            })
            continue
        parts = [p.strip() for p in el.text.split("\n") if p.strip()]
        if not parts:
            continue
        lh = el.h / len(parts)
        for j, t in enumerate(parts):
            x0 = max(0, int(el.x))
            x1 = min(W, int(el.x + el.w))
            y0 = max(0, int(el.y + j * lh))
            y1 = min(H, int(el.y + (j + 1) * lh))
            if x1 - x0 < 4 or y1 - y0 < 4:
                continue
            crop = img_bgr[y0:y1, x0:x1]
            mask = _binarize(crop)
            # 收紧 bbox：按字形像素投影裁掉块级框的空白边距
            ys, xs = np.where(mask > 0)
            if len(xs) > 20:
                y0, y1 = y0 + int(ys.min()), y0 + int(ys.max()) + 1
                x0, x1 = x0 + int(xs.min()), x0 + int(xs.max()) + 1
                crop = img_bgr[y0:y1, x0:x1]
                mask = _binarize(crop)
            # 防幻觉/防复杂背景：环域不平滑或贴彩色图形的行改判保护
            grad90, frac_sat = _ring_stats(img_bgr, x0, y0, x1, y1)
            if frac_sat > RING_SAT_FRAC_MAX or grad90 > RING_GRAD_MAX:
                protected.append({
                    "text": t[:40],
                    "reason": "llm:body but complex background",
                })
                continue
            color = _hex_to_rgb(el.color_hex) or _text_color(crop, mask)
            size_pt = max(7.0, round((y1 - y0) * px_to_pt / 0.86, 1))
            ln = TextLine(
                text=t, score=1.0,
                x=x0, y=y0, w=x1 - x0, h=y1 - y0,
                size_pt=size_pt, color=color,
                bold=el.bold or _is_bold(mask, y1 - y0),
                mask=mask,
            )
            rebuild.append(replace(ln, size_pt=_fit_size_pt(ln, dpi)))
    return rebuild, protected


# --------------------------------------------------------------------------
# Main entry
# --------------------------------------------------------------------------
def convert_pdf_to_pptx(pdf_path, pptx_path, dpi=150, manifest_path=None,
                        workdir=None, keep_workdir=False, mode="auto"):
    """mode: auto (LLM if key present, else OCR) | llm | ocr"""
    doc = pymupdf.open(pdf_path)
    if doc.page_count == 0:
        raise ValueError("PDF has no pages")

    llm = None
    if mode in ("auto", "llm"):
        from .llm import LlmLayoutExtractor
        if LlmLayoutExtractor.available():
            llm = LlmLayoutExtractor()

    # 1. render all pages
    rasters = []
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        rasters.append((cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR),
                        page.rect.width, page.rect.height))
    doc.close()

    # 2. parallel LLM extraction (falls back to OCR per page on failure)
    llm_results = [None] * len(rasters)
    if llm is not None:
        from concurrent.futures import ThreadPoolExecutor
        def _one(i):
            try:
                return i, llm.extract(rasters[i][0])
            except Exception:
                return i, None
        with ThreadPoolExecutor(max_workers=4) as ex:
            for i, els in ex.map(_one, range(len(rasters))):
                llm_results[i] = els

    ocr = None
    pages = []
    for idx, (bgr, w_pt, h_pt) in enumerate(rasters):
        source = "ocr(RapidOCR)"
        elements = llm_results[idx]
        if elements and any(e.kind == "body" for e in elements):
            rebuild, protected = lines_from_llm(bgr, elements, dpi)
            source = f"llm({llm.model})"
        else:
            if ocr is None:
                ocr = OcrEngine()
            lines = analyze_lines(bgr, ocr.lines(bgr), dpi)
            rebuild, protected = classify_lines(bgr, lines, dpi, bgr.shape[0])
        fixed = erase_lines(bgr, rebuild)
        ok, buf = cv2.imencode(".jpg", fixed,
                               [int(cv2.IMWRITE_JPEG_QUALITY), 88])
        pages.append({
            "bg_bytes": buf.tobytes(),
            "lines": rebuild,
            "protected": protected,
            "source": source,
            "w_pt": w_pt,
            "h_pt": h_pt,
        })
    out = build_pptx(pages, pptx_path, dpi)
    if manifest_path:
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
    return out
