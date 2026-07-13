/* translation-store.js
 * Per-shape dual-language cache, keyed by shape.id, stored in
 * Office.context.document.settings (shared between the task pane and
 * the headless commands.js — see manifest.xml). Each shape entry maps
 * a language code to the exact text + font last used for that
 * language, so switching direction is a cache read instead of a
 * fresh translation call, and restores the original font exactly.
 */

const TRANSLATION_STORE_KEY = "shapeTranslations";

function getTranslationStore() {
  return Office.context.document.settings.get(TRANSLATION_STORE_KEY) || {};
}

function getShapeLanguageVersion(shapeId, lang) {
  const store = getTranslationStore();
  return (store[shapeId] && store[shapeId][lang]) || null;
}

function saveShapeLanguageVersion(shapeId, lang, text, fontName) {
  const store = getTranslationStore();
  if (!store[shapeId]) {
    store[shapeId] = {};
  }
  store[shapeId][lang] = { text, font: fontName };
  Office.context.document.settings.set(TRANSLATION_STORE_KEY, store);
  return Office.context.document.settings.saveAsync();
}
