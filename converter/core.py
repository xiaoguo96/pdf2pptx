# -*- coding: utf-8 -*-
"""PDF -> editable PPTX core pipeline.

Strategy
--------
Image-only PDFs (scanned / exported-as-picture slides) are converted by:

1. Rendering each page to a high-resolution raster.
2. Extracting text elements (dual mode):
   - LLM mode: an OpenAI-compatible vision model reads the whole page and
     returns text blocks with bbox / color / bold / semantic kind;
   - OCR mode: local RapidOCR per-line recognition (offline fallback).
3. Nearly ALL text is rebuilt as native editable TextBoxes — including
   page numbers, header/footer text (company name, slogan) and TOC page
   numbers. Only elements that carry no real ink (anti-hallucination) or
   are too small / low-confidence stay untouched.
4. Background is restored per line depending on local complexity:
   - smooth background  -> row-wise linear-interpolation fill (no shadow);
   - complex / colored background (cover art, gradients, icons) ->
     tight-mask TELEA inpaint limited to the glyph pixels.
5. Font size / letter-spacing are fitted with REAL font metrics
   (Microsoft YaHei via PIL): size comes from glyph height; if the
   rendered line is narrower than the original it is letter-spaced to
   fill the exact original width, otherwise the size is shrunk so the
   text never overflows its original box.
6. Building a PPTX slide whose size equals the PDF page size:
   - the cleaned raster is inserted as one full-bleed Picture object;
   - every rebuilt line becomes a native editable TextBox.
"""

import io
import json
import os
import re
from dataclasses import dataclass, field, replace

import cv2
import numpy as np
import pymupdf
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn

EMU_PER_PT = 12700
CJK_FONT = "Microsoft YaHei"
LATIN_FONT = "Microsoft YaHei"

# side-band background complexity thresholds
RING_GRAD_MAX = 30.0      # 90th-percentile Sobel magnitude -> "smooth"
RING_SAT_FRAC_MAX = 0.30  # fraction of strongly colored band pixels


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
    erase_method: str = "row"     # "row" (smooth bg) | "telea" (complex bg)
    spacing_pt: float = 0.0       # letter spacing applied in PPTX


def _bg_color(img_bgr, x0, y0, x1, y1):
    """Median BGR of the thin side bands just outside a bbox (= background)."""
    H, W = img_bgr.shape[:2]
    y0c, y1c = max(0, y0), min(H, y1)
    x0c, x1c = max(0, x0), min(W, x1)
    parts = []
    la, lb = max(0, x0c - 16), max(0, x0c - 4)
    if lb - la >= 2:
        parts.append(img_bgr[y0c:y1c, la:lb].reshape(-1, 3))
    ra, rb = min(W, x1c + 4), min(W, x1c + 16)
    if rb - ra >= 2:
        parts.append(img_bgr[y0c:y1c, ra:rb].reshape(-1, 3))
    ta, tb = max(0, y0c - 10), max(0, y0c - 3)
    if tb - ta >= 2:
        parts.append(img_bgr[ta:tb, x0c:x1c].reshape(-1, 3))
    ba, bb = min(H, y1c + 3), min(H, y1c + 10)
    if bb - ba >= 2:
        parts.append(img_bgr[ba:bb, x0c:x1c].reshape(-1, 3))
    if not parts:
        return None
    return np.median(np.concatenate(parts, axis=0), axis=0)


def _binarize(crop_bgr, bg=None):
    """Return mask where mask>0 selects TEXT pixels.

    When the surrounding background color `bg` is known, text = pixels
    deviating from it (per-channel max distance > 30). This polarity-free
    rule replaces the old "minority pixels are text" heuristic, which
    inverted on thick large-font glyphs (tight bbox -> background became
    the minority -> background got erased -> black blobs).
    """
    if bg is not None:
        dev = np.abs(crop_bgr.astype(np.int16)
                     - np.asarray(bg, np.int16).reshape(1, 1, 3)).max(axis=2)
        return (dev > 30).astype(np.uint8) * 255
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    ret, m = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # polarity by border ring: the outermost pixels are mostly background
    border = np.concatenate([gray[0, :], gray[-1, :], gray[:, 0], gray[:, -1]])
    if np.median(border) >= ret:
        # bright side is background -> ink is the dark side (m == 0)
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
        bg = _bg_color(img_bgr, x0, y0, x1, y1)
        mask = _binarize(crop, bg)
        h_px = y1 - y0
        out.append(TextLine(
            text=ln["text"],
            score=ln["score"],
            x=x0, y=y0, w=x1 - x0, h=y1 - y0,
            size_pt=max(7.0, round(h_px * px_to_pt / 0.86, 1)),
            color=_text_color(crop, mask),
            bold=_is_bold(mask, h_px),
            mask=mask,
        ))
    return out


# --------------------------------------------------------------------------
# Font-metric based size / spacing fitting (PIL + Microsoft YaHei)
# --------------------------------------------------------------------------
_FONT_PATH = None
for _p in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf",
           r"C:\Windows\Fonts\simhei.ttf",
           "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"):
    if os.path.exists(_p):
        _FONT_PATH = _p
        break

_font_cache = {}


def _pil_font(size_px):
    key = max(8, int(round(size_px * 2)))  # 0.5 px grid
    f = _font_cache.get(key)
    if f is None:
        try:
            f = ImageFont.truetype(_FONT_PATH, key / 2.0) if _FONT_PATH \
                else ImageFont.load_default()
        except Exception:
            f = ImageFont.load_default()
        _font_cache[key] = f
    return f


def _measure_px(text, size_pt, dpi):
    """Rendered advance width of `text` at `size_pt`, in render pixels."""
    f = _pil_font(size_pt * dpi / 72.0)
    try:
        return float(f.getlength(text))
    except AttributeError:
        return float(f.getsize(text)[0])


def _fit_typo(ln, dpi):
    """Fit (size_pt, spacing_pt) with real font metrics.

    - size starts from glyph height (tight box ~= 0.86 em);
    - if the rendered line is narrower than the original ink width,
      positive letter-spacing spreads it to fill the exact width;
    - if wider, the size is shrunk so text never overflows the box.
    """
    px_per_pt = dpi / 72.0
    size_h = max(7.0, ln.h / px_per_pt / 0.86)
    w_target = float(ln.w)
    w_at_h = _measure_px(ln.text, size_h, dpi)
    if w_at_h <= 1e-6:
        return round(size_h, 1), 0.0
    if w_at_h <= w_target:
        spacing = 0.0
        n = len(ln.text)
        if n > 1 and w_target - w_at_h > 1.0:
            spacing = ((w_target - w_at_h) / (n - 1)) / px_per_pt
            spacing = max(0.0, min(spacing, 0.4 * size_h))
        return round(size_h, 1), round(spacing, 2)
    size_w = size_h * (w_target / w_at_h)
    w2 = _measure_px(ln.text, size_w, dpi)
    if w2 > 1e-6:
        size_w *= w_target / w2
    return max(7.0, round(size_w, 1)), 0.0


# --------------------------------------------------------------------------
# Line classification / segmentation (OCR path)
# --------------------------------------------------------------------------
RE_LEAD_DIGITS = re.compile(r"^(\d{2,3})(?=\D)")
RE_TRAIL_DIGITS = re.compile(r"(\d{1,3}(?:\.\d+)?%?)$")


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


def _split_segments(img, ln, dpi):
    """Split '01Title03'-style merged OCR lines into separate TextLines.

    Digit segments (page numbers / TOC numbers) are emitted as ordinary
    candidates too, so they become editable instead of protected.
    """
    t = ln.text.strip()
    lead = RE_LEAD_DIGITS.match(t)
    trail = RE_TRAIL_DIGITS.search(t)
    i0 = len(lead.group(1)) if lead else 0
    i1 = trail.start() if (trail and trail.start() >= max(3, i0 + 2)) else len(t)
    if i0 == 0 and i1 == len(t):
        return [ln]
    widths = [_char_w(c) for c in t]
    total = sum(widths) or 1.0
    H, W = img.shape[:2]
    out = []
    for a, b in ((0, i0), (i0, i1), (i1, len(t))):
        if b <= a:
            continue
        xa = ln.x + ln.w * sum(widths[:a]) / total
        xb = ln.x + ln.w * sum(widths[:b]) / total
        if xb - xa < 4:
            continue
        x0 = max(0, int(np.floor(xa)))
        x1 = min(W, int(np.ceil(xb)))
        y0 = max(0, int(ln.y))
        y1 = min(H, int(ln.y + ln.h))
        if x1 - x0 < 4 or y1 - y0 < 4:
            continue
        crop = img[y0:y1, x0:x1]
        mask = _binarize(crop, _bg_color(img, x0, y0, x1, y1))
        ys, xs = np.where(mask > 0)
        if len(xs) > 8:  # tighten to ink
            x0, x1 = x0 + int(xs.min()), x0 + int(xs.max()) + 1
            y0, y1 = y0 + int(ys.min()), y0 + int(ys.max()) + 1
            if x1 - x0 < 3 or y1 - y0 < 3:
                continue
            crop = img[y0:y1, x0:x1]
            mask = _binarize(crop, _bg_color(img, x0, y0, x1, y1))
        out.append(TextLine(
            text=t[a:b], score=ln.score,
            x=x0, y=y0, w=x1 - x0, h=y1 - y0,
            size_pt=max(7.0, round((y1 - y0) * 72.0 / dpi / 0.86, 1)),
            color=_text_color(crop, mask),
            bold=_is_bold(mask, y1 - y0),
            mask=mask,
        ))
    return out or [ln]


def classify_lines(img, lines, dpi, page_h_px):
    """Return (rebuild_lines, protected_infos).

    Everything with real ink becomes editable; protection is only for
    tiny / low-confidence fragments.
    """
    rebuild, protected = [], []
    for ln in lines:
        t = ln.text.strip()
        if not t:
            protected.append({"text": ln.text, "reason": "empty"})
            continue
        for seg in _split_segments(img, ln, dpi):
            if seg.h < 6 or not seg.text.strip():
                protected.append({"text": seg.text, "reason": "too small"})
                continue
            if seg.score < 0.7:
                protected.append({"text": seg.text,
                                  "reason": "low OCR confidence"})
                continue
            grad90, frac_sat = _ring_stats(img, int(seg.x), int(seg.y),
                                           int(seg.x + seg.w),
                                           int(seg.y + seg.h))
            complex_bg = frac_sat > RING_SAT_FRAC_MAX or grad90 > RING_GRAD_MAX
            seg = replace(seg, erase_method="telea" if complex_bg else "row")
            size_pt, spacing_pt = _fit_typo(seg, dpi)
            rebuild.append(replace(seg, size_pt=size_pt, spacing_pt=spacing_pt))
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
    """Remove rebuilt-line pixels. Modifies a copy.

    - erase_method "row": row-wise interpolation (smooth backgrounds);
    - erase_method "telea": tight-mask TELEA inpaint (complex/colored
      backgrounds such as cover art, gradients, icons).
    """
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
        if ln.erase_method == "telea":
            pad = 8
            ry0, rx0 = max(0, ln.y - pad), max(0, ln.x - pad)
            ry1, rx1 = min(h, py1 + pad), min(w, px1 + pad)
            region = out[ry0:ry1, rx0:rx1].copy()
            full_m = np.zeros(region.shape[:2], np.uint8)
            full_m[ln.y - ry0:py1 - ry0, ln.x - rx0:px1 - rx0] = m
            out[ry0:ry1, rx0:rx1] = cv2.inpaint(region, full_m, 3,
                                                cv2.INPAINT_TELEA)
        else:
            _fill_masked_rows(out[ln.y:py1, ln.x:px1], m)
    return out


# --------------------------------------------------------------------------
# PPTX building
# --------------------------------------------------------------------------
def _set_run_font(run, ln):
    f = run.font
    f.size = Pt(ln.size_pt)
    f.bold = bool(ln.bold)
    f.name = LATIN_FONT
    f.color.rgb = RGBColor(*ln.color)
    if ln.spacing_pt:
        f.spacing = Pt(ln.spacing_pt)
    rPr = run._r.get_or_add_rPr()
    el = rPr.find(qn("a:ea"))
    if el is None:
        el = rPr.makeelement(qn("a:ea"), {})
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
            _set_run_font(run, ln)
            entry["texts"].append({
                "text": ln.text, "score": ln.score,
                "x_pt": round(ln.x * 72 / dpi, 2), "y_pt": round(ln.y * 72 / dpi, 2),
                "w_pt": round(ln.w * 72 / dpi, 2), "h_pt": round(ln.h * 72 / dpi, 2),
                "size_pt": ln.size_pt, "color": list(ln.color), "bold": ln.bold,
                "spacing_pt": ln.spacing_pt, "erase": ln.erase_method,
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


def _llm_reread(llm, crop_bgr):
    """Focused re-read of one line crop by the vision model (tie-breaker)."""
    import base64
    h = crop_bgr.shape[0]
    if h < 48:
        k = 48.0 / h
        crop_bgr = cv2.resize(crop_bgr, None, fx=k, fy=k,
                              interpolation=cv2.INTER_CUBIC)
    ok, buf = cv2.imencode(".jpg", crop_bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 92])
    if not ok:
        return ""
    b64 = base64.b64encode(buf.tobytes()).decode()
    payload = {
        "model": llm.model,
        "max_tokens": 200,
        "messages": [{"role": "user", "content": [
            {"type": "text",
             "text": "请准确读出图中这一行文字，只输出文字本身，不要任何说明。"},
            {"type": "image_url",
             "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}]}],
    }
    try:
        data = llm._post(payload)
        return data["choices"][0]["message"]["content"].strip()
    except Exception:
        return ""


def _ocr_crosscheck(ocr, llm, img_bgr, x0, y0, x1, y1, llm_text):
    """Cross-check LLM-read text with local OCR on the line crop.

    Fixes occasional vision-model misreads (e.g. 柯幻 -> 构幻): when OCR is
    confident, returns a comparable-length result and clearly disagrees
    with the LLM text, the line crop is re-read by the vision model and
    the best reading wins. Fragmentary OCR results are ignored.
    """
    if ocr is None:
        return llm_text
    H, W = img_bgr.shape[:2]
    bw, bh = x1 - x0, y1 - y0
    px = max(10, int(0.12 * bw))   # generous padding: OCR needs context
    py = max(8, int(0.35 * bh))
    ya, yb = max(0, y0 - py), min(H, y1 + py)
    xa, xb = max(0, x0 - px), min(W, x1 + px)
    crop = img_bgr[ya:yb, xa:xb]
    ch = crop.shape[0]
    if crop.size == 0 or ch < 6:
        return llm_text
    ocr_crop = crop
    if ch < 28:  # upscale tiny crops so OCR can read them
        k = 28.0 / ch
        ocr_crop = cv2.resize(crop, None, fx=k, fy=k,
                              interpolation=cv2.INTER_CUBIC)
    try:
        res, _ = ocr._ocr(ocr_crop)
    except Exception:
        return llm_text
    if not res:
        return llm_text
    # join all detected pieces left-to-right (a line may split into parts)
    pieces = sorted(res, key=lambda r: min(p[0] for p in r[0]))
    otext = "".join(str(r[1]).strip() for r in pieces)
    conf = min(float(r[2]) for r in pieces)
    if conf < 0.85 or not otext:
        return llm_text
    if len(otext) < 0.8 * len(llm_text):
        return llm_text  # fragmentary OCR -> keep LLM text
    if len(otext) != len(llm_text):
        # length mismatch: arbitrate with the ink glyph count of the crop
        tight = img_bgr[y0:y1, x0:x1]
        if tight.size:
            tm = _binarize(tight, _bg_color(img_bgr, x0, y0, x1, y1))
            k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
            n, _ = cv2.connectedComponents(cv2.dilate(tm, k, iterations=1))
            glyphs = max(0, n - 1)
            if abs(len(llm_text) - glyphs) <= abs(len(otext) - glyphs):
                return llm_text
    from difflib import SequenceMatcher
    if SequenceMatcher(None, llm_text, otext).ratio() >= 0.95:
        return llm_text  # essentially identical
    # real dispute: re-read the crop and take the best of three readings
    if llm is not None:
        reread = _llm_reread(llm, crop)
        if reread:
            r_llm = SequenceMatcher(None, reread, llm_text).ratio()
            r_ocr = SequenceMatcher(None, reread, otext).ratio()
            if 0.8 * len(llm_text) <= len(reread) <= 1.25 * len(llm_text) \
                    and max(r_llm, r_ocr) >= 0.6:
                return reread
            return llm_text if r_llm >= r_ocr else otext
    return otext


def lines_from_llm(img_bgr, elements, dpi, ocr=None, llm=None):
    """Convert LLM elements to (rebuild_lines, protected_infos).

    All semantic kinds (body / header / footer / page_number / logo /
    chart_label / badge) are rebuilt as editable text boxes. Protection
    only happens when the claimed region contains no real ink (model
    hallucination guard) or is too small.
    """
    rebuild, protected = [], []
    H, W = img_bgr.shape[:2]
    for el in elements:
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
                protected.append({"text": t[:40], "reason": "llm:bbox too small"})
                continue
            crop = img_bgr[y0:y1, x0:x1]
            mask = _binarize(crop, _bg_color(img_bgr, x0, y0, x1, y1))
            ys, xs = np.where(mask > 0)
            if len(xs) < 20:
                protected.append({"text": t[:40],
                                  "reason": f"llm:{el.kind} no ink"})
                continue
            # tighten bbox to ink projection
            y0, y1 = y0 + int(ys.min()), y0 + int(ys.max()) + 1
            x0, x1 = x0 + int(xs.min()), x0 + int(xs.max()) + 1
            crop = img_bgr[y0:y1, x0:x1]
            mask = _binarize(crop, _bg_color(img_bgr, x0, y0, x1, y1))
            ink_frac = float((mask > 0).mean())
            # anti-hallucination: claimed text over complex backgrounds must
            # contain a credible amount of ink
            grad90, frac_sat = _ring_stats(img_bgr, x0, y0, x1, y1)
            complex_bg = frac_sat > RING_SAT_FRAC_MAX or grad90 > RING_GRAD_MAX
            if ink_frac < (0.03 if complex_bg else 0.01):
                protected.append({"text": t[:40],
                                  "reason": f"llm:{el.kind} ink too sparse"})
                continue
            if el.kind in ("chart_label", "badge") and complex_bg:
                # labels sitting on chart graphics (lines/bars/badges):
                # erasing them would damage the graphic -> keep original pixels
                protected.append({"text": t[:40],
                                  "reason": f"llm:{el.kind} on chart graphic"})
                continue
            t = _ocr_crosscheck(ocr, llm, img_bgr, x0, y0, x1, y1, t)
            if el.kind == "page_number" and llm is not None:
                # page numbers are read most reliably from a zoomed crop
                rr = _llm_reread(llm, img_bgr[y0:y1, x0:x1]).strip()
                if rr.isdigit():
                    t = rr
            color = _hex_to_rgb(el.color_hex) or _text_color(crop, mask)
            ln = TextLine(
                text=t, score=1.0,
                x=x0, y=y0, w=x1 - x0, h=y1 - y0,
                size_pt=0.0, color=color,
                bold=el.bold or _is_bold(mask, y1 - y0),
                mask=mask,
                erase_method="telea" if complex_bg else "row",
            )
            size_pt, spacing_pt = _fit_typo(ln, dpi)
            rebuild.append(replace(ln, size_pt=size_pt, spacing_pt=spacing_pt))
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
        if elements:
            if ocr is None:
                ocr = OcrEngine()
            rebuild, protected = lines_from_llm(bgr, elements, dpi, ocr, llm)
            if not rebuild:  # model returned nothing usable -> OCR fallback
                elements = None
            else:
                source = f"llm({llm.model})"
        if not elements:
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
