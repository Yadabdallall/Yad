/**
 * پۆرتاڵی چالاکییەکانی بۆردی وەرزش — کۆکردنەوەی خۆکاری فۆڕمەکان
 *
 * هەر فۆڕمێک کە دەنێردرێت لێرە لە Cloudflare KV هەڵدەگیرێت، و بەشی ئیدارە
 * ئامارەکان ڕاستەوخۆ لێرەوە وەردەگرێت.
 *
 *   POST /api/submit   فۆڕمێک هەڵدەگرێت (هەمووان دەتوانن بنێرن)
 *   GET  /api/stats    هەموو فۆڕمەکان (تەنها بە کۆدی ئیدارە، سەرپەڕەی x-admin-key)
 *
 * پێویستی بە یەک KV binding هەیە بە ناوی FORMS.
 * ئەم فایلە لەناو Cloudflare Pages ـدا کار دەکات (_worker.js)، یان وەک Worker ـێکی
 * سەربەخۆ؛ لە حاڵەتی دووەمدا API_BASE لە index.html دەگۆڕدرێت بۆ لینکی Worker ـەکە.
 */

// تەنها SHA-256 ی کۆدی ئیدارە، نەک خودی کۆدەکە
const ADMIN_HASH = "244022060493ff3467dfe5248b844a2094b1a13c83482e4eba8c977a8d11310d";
const MAX_BODY = 20000;

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "content-type, x-admin-key",
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store", ...CORS },
  });
}

async function sha256(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

const str = (v, max) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, max);
const int = (v) => {
  const n = parseInt(String(v ?? "").replace(/[^\d]/g, ""), 10);
  return Number.isFinite(n) && n >= 0 && n < 1e13 ? n : 0;
};

// ئەو زانیارییانەی ئامار پێویستی پێیانە — لە metadata ـی KV ـدا (کەمتر لە ١٠٢٤ بایت)
function compact(d, t) {
  const row = {
    t, ref: d.ref,
    rep: str(d.repName, 40), city: str(d.city, 40), unit: str(d.unit, 50),
    act: str(d.actName, 60), type: str(d.actType, 40), date: str(d.actDate, 10),
    n: int(d.nAll), w: int(d.nWomen), y: int(d.nYouth),
    cost: int(d.cost), cur: d.currency === "USD" ? "USD" : "IQD", fund: str(d.fund, 40),
  };
  while (new TextEncoder().encode(JSON.stringify(row)).length > 1000 && row.act.length > 10) {
    row.act = row.act.slice(0, -10); row.unit = row.unit.slice(0, -5);
  }
  return row;
}

async function submit(request, env) {
  const text = await request.text();
  if (text.length > MAX_BODY) return json({ ok: false, error: "too-large" }, 413);
  let body;
  try { body = JSON.parse(text); } catch (e) { return json({ ok: false, error: "bad-json" }, 400); }
  if (body.hp) return json({ ok: true });                       // خانەی شاراوە پڕکراوەتەوە: بۆت
  const d = body.data || {};
  if (!/^SB-\d{6}-\d{4}$/.test(String(d.ref || ""))) return json({ ok: false, error: "bad-ref" }, 400);
  if (!str(d.repName, 1) || !str(d.city, 1) || !str(d.actName, 1)) return json({ ok: false, error: "missing" }, 400);

  const key = "f:" + d.ref;
  if (await env.FORMS.get(key)) return json({ ok: true, duplicate: true });
  const t = new Date().toISOString();
  const full = { t, ...Object.fromEntries(Object.entries(d).map(([k, v]) => [k, typeof v === "string" ? v.slice(0, 5000) : v])) };
  await env.FORMS.put(key, JSON.stringify(full), { metadata: compact(d, t) });
  return json({ ok: true, ref: d.ref });
}

async function stats(request, env) {
  const key = request.headers.get("x-admin-key") || "";
  if ((await sha256(key)) !== ADMIN_HASH) return json({ ok: false, error: "auth" }, 401);
  const rows = [];
  let cursor;
  do {
    const page = await env.FORMS.list({ prefix: "f:", cursor });
    for (const k of page.keys) {
      if (k.metadata) rows.push(k.metadata);
      else {
        const v = await env.FORMS.get(k.name, "json");
        if (v) rows.push(compact(v, v.t));
      }
    }
    cursor = page.list_complete ? undefined : page.cursor;
  } while (cursor);
  return json({ ok: true, rows });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname.startsWith("/api/")) {
      if (request.method === "OPTIONS") return new Response(null, { headers: CORS });
      if (!env.FORMS) return json({ ok: false, error: "no-storage" }, 503);
      try {
        if (url.pathname === "/api/submit" && request.method === "POST") return await submit(request, env);
        if (url.pathname === "/api/stats" && request.method === "GET") return await stats(request, env);
      } catch (err) {
        return json({ ok: false, error: "server" }, 500);
      }
      return json({ ok: false, error: "not-found" }, 404);
    }
    // هەموو شتێکی تر: پەڕەکانی ماڵپەڕەکە
    return env.ASSETS ? env.ASSETS.fetch(request) : new Response("Not found", { status: 404 });
  },
};
