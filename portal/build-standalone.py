#!/usr/bin/env python3
"""Build a single self-contained HTML file from portal/index.html.

Fonts are inlined as data URIs and jsPDF is inlined as a script, so the file
works on its own — sent by WhatsApp, opened from Files, or uploaded anywhere.
"""
import base64, pathlib, re

here = pathlib.Path(__file__).parent
html = (here / "index.html").read_text(encoding="utf-8")

def font_uri(m):
    data = (here / m.group(1)).read_bytes()
    return "url('data:font/woff2;base64," + base64.b64encode(data).decode() + "')"

html = re.sub(r"url\('(assets/fonts/[^']+\.woff2)'\)", font_uri, html)

def img_uri(m):
    path = m.group(1)
    mime = "image/jpeg" if path.endswith((".jpg", ".jpeg")) else "image/png" if path.endswith(".png") else "image/webp"
    return 'src="data:' + mime + ';base64,' + base64.b64encode((here / path).read_bytes()).decode() + '"'

html = re.sub(r'src="(assets/img/[^"]+)"', img_uri, html)

tag = '<script src="assets/js/jspdf.umd.min.js" defer></script>'
assert tag in html
js = (here / "assets/js/jspdf.umd.min.js").read_text(encoding="utf-8")
html = html.replace(tag + "\n", "").replace(tag, "")
html = html.replace("</body>", "<script>\n" + js + "\n</script>\n</body>", 1)

out = here / "PUK SPORTS BOARD - PORTAL.html"
out.write_text(html, encoding="utf-8")
print(f"built {out.name} ({out.stat().st_size // 1024} KB)")
