/* translate-client.js
 * Shared by taskpane.js and commands.js. Calls the local backend proxy
 * (server/index.js), which forwards to Google Cloud Translation.
 * Loaded as a plain script (not a module) so both html pages can include
 * it with a simple <script> tag ahead of their own logic.
 */

async function callTranslateApi(text, { sourceLang, targetLang }) {
  const response = await fetch("https://localhost:3000/translate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text, sourceLang, targetLang }),
  });

  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(data.error || `Translate API error: ${response.status}`);
  }

  return data.translated;
}

/* translateSelectedShape
 * Shared by taskpane.js and commands.js: translates the shape currently
 * selected in PowerPoint, in place, using the given language pair.
 *
 * Checks the per-shape dual-language cache (translation-store.js) first —
 * if this shape already has a cached version in targetLang, that exact
 * text + font is restored with no API call. Otherwise it calls the
 * translate API, applies a CJK-appropriate font when translating into
 * Chinese (font-map.js), and caches both the source and target versions
 * for instant toggling next time.
 */
async function translateSelectedShape({ sourceLang, targetLang }) {
  await PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    await context.sync();

    if (selectedShapes.items.length === 0) {
      throw new Error("No shape selected.");
    }

    const shape = selectedShapes.items[0];
    shape.load("id");
    const textRange = shape.textFrame.textRange;
    textRange.load("text");
    const font = textRange.font;
    font.load("name");
    await context.sync();

    const originalText = textRange.text;
    if (!originalText || !originalText.trim()) {
      throw new Error("Selected shape has no text.");
    }

    const shapeId = shape.id;

    // Always remember what's currently on the slide under its own language,
    // so the first translation in either direction has something to restore to.
    if (!getShapeLanguageVersion(shapeId, sourceLang)) {
      await saveShapeLanguageVersion(shapeId, sourceLang, originalText, font.name);
    }

    const cached = getShapeLanguageVersion(shapeId, targetLang);

    let translatedText;
    let targetFont;
    if (cached) {
      translatedText = cached.text;
      targetFont = cached.font;
    } else {
      translatedText = await callTranslateApi(originalText, { sourceLang, targetLang });
      targetFont = pickFontForTranslation(targetLang, font.name);
      await saveShapeLanguageVersion(shapeId, targetLang, translatedText, targetFont);
    }

    textRange.text = translatedText;
    if (targetFont) {
      font.name = targetFont;
    }
    await context.sync();
  });
}
