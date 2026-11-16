# -*- coding: utf-8 -*-
"""CLI: python scripts/convert.py input.pdf output.pptx [--dpi 150]"""
import argparse
import collections
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from converter import convert_pdf_to_pptx


def main():
    ap = argparse.ArgumentParser(description="Convert image-based PDF to editable PPTX")
    ap.add_argument("pdf")
    ap.add_argument("pptx")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--manifest", default=None,
                    help="optional path to write element manifest JSON")
    ap.add_argument("--mode", choices=["auto", "llm", "ocr"], default="auto",
                    help="auto: 有 LLM_API_KEY 时用视觉大模型，否则本地 OCR")
    args = ap.parse_args()
    manifest = args.manifest or (os.path.splitext(args.pptx)[0] + ".manifest.json")
    m = convert_pdf_to_pptx(args.pdf, args.pptx, dpi=args.dpi,
                            manifest_path=manifest, mode=args.mode)
    n_txt = sum(len(p["texts"]) for p in m)
    n_prot = sum(len(p.get("protected", [])) for p in m)
    src = collections.Counter(p.get("source", "?") for p in m)
    print(f"OK: {len(m)} slides, {n_txt} editable text objects, "
          f"{n_prot} protected elements -> {args.pptx}")
    print("   sources:", dict(src))


if __name__ == "__main__":
    main()
