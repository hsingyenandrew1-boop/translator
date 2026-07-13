/* server/index.js
 * Local HTTPS server: serves the add-in's static files (manifest targets
 * https://localhost:3000/...) and proxies /translate to Google Cloud
 * Translation, keeping the API key off the client.
 */

require("dotenv").config();
const express = require("express");
const path = require("path");
const https = require("https");
const devCerts = require("office-addin-dev-certs");

const PORT = process.env.PORT || 3000;
const GOOGLE_TRANSLATE_API_KEY = process.env.GOOGLE_TRANSLATE_API_KEY;

if (!GOOGLE_TRANSLATE_API_KEY) {
  console.warn("Warning: GOOGLE_TRANSLATE_API_KEY is not set — /translate will fail until it is (see .env.example).");
}

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "..")));

app.post("/translate", async (req, res) => {
  const { text, sourceLang, targetLang } = req.body || {};

  if (!text || !targetLang) {
    return res.status(400).json({ error: "text and targetLang are required" });
  }

  try {
    const params = new URLSearchParams({
      q: text,
      target: targetLang,
      format: "text", // otherwise Google HTML-escapes punctuation, which PPT text doesn't expect
      key: GOOGLE_TRANSLATE_API_KEY,
    });
    if (sourceLang) params.set("source", sourceLang);

    const response = await fetch(`https://translation.googleapis.com/language/translate/v2?${params}`, {
      method: "POST",
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error?.message || `Google Translate API error: ${response.status}`);
    }

    res.json({ translated: data.data.translations[0].translatedText });
  } catch (err) {
    console.error("Translate request failed:", err);
    res.status(502).json({ error: err.message });
  }
});

devCerts.getHttpsServerOptions().then((httpsOptions) => {
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`Bilingual PPT Mode server running at https://localhost:${PORT}`);
  });
});
