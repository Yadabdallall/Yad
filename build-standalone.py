#!/usr/bin/env python3
"""دروستکردنی وەشانی تاکە-فایل: هەموو وێنە و فۆنتەکان دەخاتە ناو index.html"""
import base64, io, os, re

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "wlat-sports-form-standalone.html")
MIME = {".woff2":"font/woff2", ".webp":"image/webp", ".png":"image/png",
        ".jpg":"image/jpeg", ".ttf":"font/ttf"}

html = io.open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()

for path in sorted(set(re.findall(r"assets/[\w./-]+", html)), key=len, reverse=True):
    full = os.path.join(ROOT, path)
    if not os.path.exists(full):
        raise SystemExit("missing asset: " + path)
    mime = MIME[os.path.splitext(path)[1]]
    data = base64.b64encode(open(full, "rb").read()).decode()
    html = html.replace(path, "data:%s;base64,%s" % (mime, data))

io.open(OUT, "w", encoding="utf-8").write(html)
print("built %s (%d KB)" % (os.path.basename(OUT), os.path.getsize(OUT) / 1024))
