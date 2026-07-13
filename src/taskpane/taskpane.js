/* taskpane.js
 * UI for the "Translate Panel" task pane: pick a language direction and
 * translate the currently selected shape. Shares document.settings with
 * commands.js so the Quick Translate ribbon button uses the same
 * direction chosen here.
 */

Office.onReady(() => {
  loadStoredLanguages();
  showLastErrorIfAny();

  document.getElementById("source-lang").addEventListener("change", saveLanguages);
  document.getElementById("target-lang").addEventListener("change", saveLanguages);
  document.getElementById("translate-button").addEventListener("click", translateSelection);
});

function loadStoredLanguages() {
  const settings = Office.context.document.settings;
  document.getElementById("source-lang").value = settings.get("sourceLang") || "en";
  document.getElementById("target-lang").value = settings.get("targetLang") || "zh-TW";
}

function saveLanguages() {
  const settings = Office.context.document.settings;
  settings.set("sourceLang", document.getElementById("source-lang").value);
  settings.set("targetLang", document.getElementById("target-lang").value);
  settings.saveAsync();
}

function showLastErrorIfAny() {
  const settings = Office.context.document.settings;
  const lastError = settings.get("lastError");
  if (lastError) {
    showBanner(`Quick Translate failed: ${lastError.message}`);
    settings.remove("lastError");
    settings.saveAsync();
  }
}

function showBanner(message) {
  const banner = document.getElementById("error-banner");
  banner.textContent = message;
  banner.hidden = false;
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

async function translateSelection() {
  setStatus("Translating...");
  try {
    const sourceLang = document.getElementById("source-lang").value;
    const targetLang = document.getElementById("target-lang").value;
    await translateSelectedShape({ sourceLang, targetLang });
    setStatus("Done.");
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message}`);
  }
}
