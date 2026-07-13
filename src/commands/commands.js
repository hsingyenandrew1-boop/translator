/* commands.js
 * Runs the "Quick Translate" ribbon button action defined in manifest.xml:
 * translate the selected shape's text in place, formatting untouched.
 *
 * Language direction is read from Office.context.document.settings, which
 * is shared with the task pane (see taskpane.js) since both are bound to
 * the same presentation document rather than to a browser storage origin.
 */

Office.onReady(() => {
  // Office.js requires actions to be associated once the runtime is ready.
});

async function quickTranslateSelection(event) {
  try {
    const { sourceLang, targetLang } = getStoredLanguages();
    await translateSelectedShape({ sourceLang, targetLang });
    clearLastError();
  } catch (err) {
    // Ribbon-triggered functions run headless — there's no UI here to show
    // an alert in, so the error is written to document settings and the
    // task pane surfaces it next time it's opened.
    console.error("quickTranslateSelection failed:", err);
    recordLastError(err.message);
  } finally {
    // Required: tells Office this function-based command has finished.
    event.completed();
  }
}

function getStoredLanguages() {
  const settings = Office.context.document.settings;
  return {
    sourceLang: settings.get("sourceLang") || "en",
    targetLang: settings.get("targetLang") || "zh-TW",
  };
}

function recordLastError(message) {
  const settings = Office.context.document.settings;
  settings.set("lastError", { message, at: Date.now() });
  settings.saveAsync();
}

function clearLastError() {
  const settings = Office.context.document.settings;
  settings.remove("lastError");
  settings.saveAsync();
}

// Register the function so the manifest's <ExecuteFunction> action can find it.
Office.actions.associate("quickTranslateSelection", quickTranslateSelection);
