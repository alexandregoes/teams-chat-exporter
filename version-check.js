// Shared by the popup (updater.js) and the service worker (background.js):
// where the repository lives, how to read its published version and where the result is stored.
const TceUpdate = (() => {
  const REPO = "alexandregoes/teams-chat-exporter";
  const REPO_URL = `https://github.com/${REPO}`;
  const REMOTE_MANIFEST_URL = `https://raw.githubusercontent.com/${REPO}/main/manifest.json`;
  const ZIP_URL = `${REPO_URL}/archive/refs/heads/main.zip`;
  const CHECK_INTERVAL_MS = 24 * 60 * 60 * 1000;
  const KEY_AUTO = "autoUpdateCheck";
  const KEY_LAST = "lastUpdateCheck";

  const hasStorage = typeof chrome !== "undefined" && chrome.storage && chrome.storage.local;

  // Returns > 0 when a is newer than b ("1.10.0" vs "1.9.3" compares numerically)
  function compareVersions(a, b) {
    const pa = String(a).split(".").map((n) => parseInt(n, 10) || 0);
    const pb = String(b).split(".").map((n) => parseInt(n, 10) || 0);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
      const diff = (pa[i] || 0) - (pb[i] || 0);
      if (diff !== 0) return diff;
    }
    return 0;
  }

  async function fetchRemoteVersion() {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 8000);
    try {
      const res = await fetch(REMOTE_MANIFEST_URL, { cache: "no-store", signal: controller.signal });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const manifest = await res.json();
      if (!manifest || typeof manifest.version !== "string") {
        throw new Error("manifest sem versão");
      }
      return manifest.version;
    } finally {
      clearTimeout(timer);
    }
  }

  // { auto: boolean, last: { at: epoch ms, version: string } | null }
  async function getState() {
    if (!hasStorage) return { auto: true, last: null };
    const data = await chrome.storage.local.get([KEY_AUTO, KEY_LAST]);
    const last = data[KEY_LAST];
    return {
      auto: data[KEY_AUTO] !== false,
      last: last && typeof last.version === "string" ? last : null
    };
  }

  async function setAuto(enabled) {
    if (hasStorage) await chrome.storage.local.set({ [KEY_AUTO]: Boolean(enabled) });
  }

  async function saveLast(version) {
    if (hasStorage) await chrome.storage.local.set({ [KEY_LAST]: { at: Date.now(), version } });
  }

  function isStale(last) {
    return !last || Date.now() - last.at >= CHECK_INTERVAL_MS;
  }

  return {
    REPO_URL,
    ZIP_URL,
    KEY_AUTO,
    KEY_LAST,
    compareVersions,
    fetchRemoteVersion,
    getState,
    setAuto,
    saveLast,
    isStale
  };
})();
