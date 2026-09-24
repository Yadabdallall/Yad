/**
 * وەرگرتنی فایلی PDF ـی فۆڕمەکە و دانانی لە R2،
 * پاشان لینکێکی گشتی دەگەڕێنێتەوە بۆ ناردن بە وەتسئەپ.
 *
 * بەستنەوەی پێویست (Bindings):
 *   R2 bucket  →  ناوی گۆڕاو: FORMS
 *
 * ڕێگاکان:
 *   POST /upload   → لەشی داواکاری = بایتەکانی PDF، وەڵام = { "url": "..." }
 *   GET  /f/<id>   → فایلەکە دەگەڕێنێتەوە
 */

const MAX_SIZE = 8 * 1024 * 1024;        // ٨ مێگابایت
const ALLOWED_ORIGIN = "*";              // بۆ توندتر کردن، ناونیشانی سایتەکەت لێرە دابنێ

function cors(extra) {
  return Object.assign({
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  }, extra || {});
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: cors() });
    }

    // ---- بەرزکردنەوە ----
    if (request.method === "POST" && url.pathname === "/upload") {
      const body = await request.arrayBuffer();

      if (!body.byteLength || body.byteLength > MAX_SIZE) {
        return Response.json({ error: "bad size" }, { status: 400, headers: cors() });
      }

      const id = crypto.randomUUID().replace(/-/g, "").slice(0, 16);
      const key = `${new Date().toISOString().slice(0, 10)}/${id}.pdf`;

      await env.FORMS.put(key, body, {
        httpMetadata: { contentType: "application/pdf" }
      });

      return Response.json(
        { url: `${url.origin}/f/${key}` },
        { headers: cors() }
      );
    }

    // ---- داگرتن ----
    if (request.method === "GET" && url.pathname.startsWith("/f/")) {
      const key = decodeURIComponent(url.pathname.slice(3));
      const object = await env.FORMS.get(key);

      if (!object) return new Response("Not found", { status: 404 });

      return new Response(object.body, {
        headers: cors({
          "Content-Type": "application/pdf",
          "Content-Disposition": 'inline; filename="PUK-Sports-Form.pdf"',
          "Cache-Control": "public, max-age=31536000, immutable"
        })
      });
    }

    return new Response("PUK Sports Board — form uploader", {
      headers: cors({ "Content-Type": "text/plain; charset=utf-8" })
    });
  }
};
