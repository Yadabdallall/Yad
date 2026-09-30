# ناردنی خۆکاری PDF بۆ وەتسئەپ

بە بنەڕەت، فۆڕمەکە PDF دروست دەکات و پەنجەرەی ناردنی ئامێرەکە دەکاتەوە.
ئەگەر دەتەوێت بە **یەک کلیک** بچێتە ژمارەی بۆرد، ئەم Worker ـە دابنێ:
PDF ـەکە بەرز دەکرێتەوە و لینکەکەی لەناو نامەی وەتسئەپدا دەنێردرێت.

## دانانی (٥ خولەک، خۆڕایی)

1. **دروستکردنی R2 bucket**
   - Cloudflare → **R2** → **Create bucket**
   - ناو: `puk-forms`

2. **دروستکردنی Worker**
   - **Workers & Pages** → **Create** → **Start with Hello World!**
   - ناو: `puk-form-upload`
   - **Deploy** → پاشان **Edit code**
   - ناوەڕۆکی `upload-worker.js` کۆپی بکە و لە جیاتی کۆدەکە دایبنێ → **Deploy**

3. **بەستنەوەی bucket ـەکە**
   - لە Worker ـەکە → **Settings** → **Bindings** → **Add** → **R2 bucket**
   - Variable name: `FORMS`
   - R2 bucket: `puk-forms`
   - **Deploy**

4. **دانانی لینکەکە لە فۆڕمەکەدا**
   - لینکی Worker ـەکە کۆپی بکە (وەک `https://puk-form-upload.xxx.workers.dev`)
   - لە `index.html` بەدوای ئەم دێڕەدا بگەڕێ و لینکەکەی تێبنووسە:

   ```js
   var UPLOAD_ENDPOINT = "";
   ```
   ↓
   ```js
   var UPLOAD_ENDPOINT = "https://puk-form-upload.xxx.workers.dev";
   ```

## دوای ئەوە چی ڕوودەدات

1. کەسەکە دوگمەکە لێدەدات
2. PDF دروست دەبێت و بەرز دەکرێتەوە (٢–٣ چرکە)
3. وەتسئەپ ڕاستەوخۆ لەگەڵ ژمارەی بۆرد دەکرێتەوە، نامەکە ئامادەیە:
   پوختەی زانیارییەکان + لینکی PDF
4. تەنها **Send** لێدەدات ✅

## تێچوو

خۆڕایی — R2 ـی Cloudflare ١٠ گیگابایت و Workers ڕۆژانە ١٠٠,٠٠٠ داواکاری بەخۆڕایی دەدات.
