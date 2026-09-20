#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
پێوانەکردنی لاپەڕەکان: DOCX → PDF → دۆزینەوەی ژمارەی پەڕەی هەر سەردێڕێک.

نیشانەکانی @@Tnnn@@ کە بە ڕەنگی سپی و قەبارەی ١ لە ناو سەردێڕەکاندان
لێرەدا لە PDF ـەکەوە دەخوێنرێنەوە بۆ دروستکردنی پێڕستێکی ڕاست.
"""
import os
import re
import sys
import json
import shutil
import subprocess

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "build")


def to_pdf(docx_path):
    workdir = os.path.join(OUT, "_pdf")
    os.makedirs(workdir, exist_ok=True)
    env = dict(os.environ, HOME=os.path.join(OUT, "_lo"))
    os.makedirs(env["HOME"], exist_ok=True)
    subprocess.run(
        ["soffice", "--headless", "--norestore", "--convert-to", "pdf",
         "--outdir", workdir, docx_path],
        check=True, capture_output=True, env=env, timeout=600,
    )
    pdf = os.path.join(
        workdir, os.path.basename(docx_path).rsplit(".", 1)[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("PDF نەدروستکرا")
    return pdf


def scan(pdf_path):
    from pypdf import PdfReader
    reader = PdfReader(pdf_path)
    pages = {}
    for n, page in enumerate(reader.pages, 1):
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        for anchor in re.findall(r"@@\s*(T\d{3})\s*@@", txt):
            pages.setdefault(anchor, n)
    return len(reader.pages), pages


def main():
    docx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        OUT, "serok_komari_yariga.docx")
    pdf = to_pdf(docx)
    total, pages = scan(pdf)
    with open(os.path.join(OUT, "toc_pages.json"), "w", encoding="utf-8") as fh:
        json.dump(pages, fh, ensure_ascii=False, indent=1)
    entries_path = os.path.join(OUT, "toc_entries.json")
    found = len(pages)
    expected = 0
    if os.path.exists(entries_path):
        with open(entries_path, encoding="utf-8") as fh:
            expected = len(json.load(fh))
    print(json.dumps({"total_pages": total, "anchors_found": found,
                      "anchors_expected": expected}, ensure_ascii=False))
    return total


if __name__ == "__main__":
    main()
