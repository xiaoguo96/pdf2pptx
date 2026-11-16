# -*- coding: utf-8 -*-
"""CLI: python scripts/convert.py input.pdf output.pptx [--dpi 150]"""
import argparse
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
    args = ap.parse_args()
    manifest = args.manifest or (os.path.splitext(args.pptx)[0] + ".manifest.json")
    m = convert_pdf_to_pptx(args.pdf, args.pptx, dpi=args.dpi,
                            manifest_path=manifest)
    n_txt = sum(len(p["texts"]) for p in m)
    print(f"OK: {len(m)} slides, {n_txt} editable text objects -> {args.pptx}")


if __name__ == "__main__":
    main()
