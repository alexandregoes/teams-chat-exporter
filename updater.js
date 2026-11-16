// "Sobre" tab of the popup: shows the installed version and whether GitHub has a newer one.
// The daily check and the toolbar badge live in background.js; shared helpers in version-check.js.
(() => {
  const versionLabel = document.getElementById("versionLabel");
  const repoLink = document.getElementById("repoLink");
  const issuesLink = document.getElementById("issuesLink");
  const autoCheckEl = document.getElementById("autoUpdateCheck");
  const checkBtn = document.getElementById("checkUpdateBtn");
  const statusEl = document.getElementById("updateStatus");
  const updateBox = document.getElementById("updateBox");
  const updateDot = document.getElementById("updateDot");
  const updateTitle = document.getElementById("updateTitle");
  const downloadLink = document.getElementById("downloadUpdateLink");
  const reloadBtn = document.getElementById("reloadExtBtn");

  const hasRuntime = typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.getManifest;
  const currentVersion = hasRuntime ? chrome.runtime.getManifest().version : "0.0.0";

  function setStatus(text, isError = false) {
    statusEl.textContent = text;
    statusEl.className = isError ? "update-status error" : "update-status";
  }

  function showResult(remoteVersion, announceUpToDate) {
    if (TceUpdate.compareVersions(remoteVersion, currentVersion) > 0) {
      updateTitle.textContent = `Nova versão ${remoteVersion} disponível`;
      updateBox.hidden = false;
      // Red dot on the "Sobre" tab, so the notice is visible from the main tab
      updateDot.hidden = false;
      setStatus("");
    } else {
      updateBox.hidden = true;
      updateDot.hidden = true;
      if (announceUpToDate) setStatus("Você está na versão mais recente.");
    }
  }

  async function checkForUpdate(manual) {
    checkBtn.disabled = true;
    if (manual) setStatus("Verificando...");
    try {
      const remoteVersion = await TceUpdate.fetchRemoteVersion();
      // Saving the result also makes the service worker refresh the toolbar badge
      await TceUpdate.saveLast(remoteVersion);
      showResult(remoteVersion, true);
    } catch {
      // A failed automatic check stays quiet; only a manual one reports the error
      if (manual) setStatus("Não foi possível consultar o GitHub.", true);
    } finally {
      checkBtn.disabled = false;
    }
  }

  versionLabel.textContent = currentVersion;
  repoLink.href = TceUpdate.REPO_URL;
  issuesLink.href = `${TceUpdate.REPO_URL}/issues`;
  downloadLink.href = TceUpdate.ZIP_URL;

  autoCheckEl.addEventListener("change", async () => {
    await TceUpdate.setAuto(autoCheckEl.checked);
    if (autoCheckEl.checked) checkForUpdate(false);
  });

  checkBtn.addEventListener("click", () => checkForUpdate(true));

  reloadBtn.addEventListener("click", () => {
    if (hasRuntime) chrome.runtime.reload();
  });

  (async function init() {
    const { auto, last } = await TceUpdate.getState();
    autoCheckEl.checked = auto;
    // Show what is already known (same information the toolbar badge uses)
    if (last) showResult(last.version, false);
    if (auto && TceUpdate.isStale(last)) checkForUpdate(false);
  })();
})();
