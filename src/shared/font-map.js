/* font-map.js
 * Latin fonts don't have full CJK glyph coverage, so translating into
 * Chinese needs a font swap to stay legible. Translating back to a
 * language we've already cached restores that language's stored font
 * exactly (see translation-store.js) — this map is only the fallback
 * for the first translation into a CJK target.
 */

const CJK_FALLBACK_FONT = {
  "zh-TW": "Microsoft JhengHei",
  "zh-CN": "Microsoft YaHei",
};

function isCjkLang(lang) {
  return Object.prototype.hasOwnProperty.call(CJK_FALLBACK_FONT, lang);
}

function pickFontForTranslation(targetLang, currentFontName) {
  if (isCjkLang(targetLang)) {
    return CJK_FALLBACK_FONT[targetLang];
  }
  return currentFontName;
}
