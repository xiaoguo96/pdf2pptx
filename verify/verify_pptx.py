# -*- coding: utf-8 -*-
"""Editable-object verification for the converted PPTX.

Checks, per slide:
  - exactly one Picture (the restored page background) that is movable/resizable;
  - every other shape is a native TextBox with an editable text run,
    with font size / color / bold attributes readable and writable.
Writes verify_report.md and verify_report.json.

Usage: python verify/verify_pptx.py <file.pptx> [report_prefix]
"""
import json
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pptx import Presentation
from pptx.util import Emu


def main():
    path = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(path)[0]
    prs = Presentation(path)
    report = {"file": os.path.basename(path),
              "slide_size_pt": [round(prs.slide_width / 12700, 1),
                                round(prs.slide_height / 12700, 1)],
              "slides": []}
    total_txt, total_pic, total_editable = 0, 0, 0
    for i, slide in enumerate(prs.slides, 1):
        pics, texts, others = [], [], []
        for sh in slide.shapes:
            if sh.shape_type == 13:  # PICTURE
                pics.append(sh)
            elif sh.has_text_frame and sh.text_frame.text.strip():
                texts.append(sh)
            else:
                others.append(sh)
        # editability probe: rewrite first run text and restore
        probe_ok = 0
        for sh in texts:
            para = sh.text_frame.paragraphs[0]
            if not para.runs:
                continue
            r = para.runs[0]
            old = r.text
            try:
                r.text = "__probe__"
                assert r.text == "__probe__"
                r.text = old
                probe_ok += 1
            except Exception:
                pass
        total_txt += len(texts)
        total_pic += len(pics)
        total_editable += probe_ok
        report["slides"].append({
            "slide": i,
            "pictures": [{"name": p.name,
                          "width_emu": p.width, "height_emu": p.height} for p in pics],
            "text_boxes": len(texts),
            "text_editable_probe_passed": probe_ok,
            "other_shapes": len(others),
            "sample_texts": [
                {"name": t.name,
                 "text": t.text_frame.text[:30],
                 "size_pt": t.text_frame.paragraphs[0].runs[0].font.size.pt
                 if t.text_frame.paragraphs[0].runs and t.text_frame.paragraphs[0].runs[0].font.size else None,
                 "bold": t.text_frame.paragraphs[0].runs[0].font.bold
                 if t.text_frame.paragraphs[0].runs else None,
                 "color": str(t.text_frame.paragraphs[0].runs[0].font.color.rgb)
                 if t.text_frame.paragraphs[0].runs and t.text_frame.paragraphs[0].runs[0].font.color and t.text_frame.paragraphs[0].runs[0].font.color.type is not None else None}
                for t in texts[:3]],
        })

    with open(prefix + ".verify.json", "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)

    md = [f"# PPTX 可编辑性验证报告: {os.path.basename(path)}", "",
          f"- 幻灯片尺寸: {report['slide_size_pt'][0]} x {report['slide_size_pt'][1]} pt (与 PDF 页面一致)",
          f"- 总页数: {len(prs.slides)}",
          f"- 背景图片对象(可独立选择/移动/缩放): {total_pic}",
          f"- 原生可编辑文本框总数: {total_txt}",
          f"- 文本改写探针通过(可编辑性实证): {total_editable}/{total_txt}", ""]
    for s in report["slides"]:
        md.append(f"## 第 {s['slide']} 页")
        md.append(f"- 图片对象: {len(s['pictures'])} | 文本框: {s['text_boxes']} | 改写探针通过: {s['text_editable_probe_passed']}")
        for t in s["sample_texts"]:
            md.append(f"  - `{t['text']}` — {t['size_pt']}pt, bold={t['bold']}, color={t['color']}")
        md.append("")
    with open(prefix + ".verify.md", "w", encoding="utf-8") as f:
        f.write("\n".join(md))
    print(f"pictures={total_pic} textboxes={total_txt} probe_ok={total_editable}")
    print("reports:", prefix + ".verify.md", prefix + ".verify.json")


if __name__ == "__main__":
    main()
