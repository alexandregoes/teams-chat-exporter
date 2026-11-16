// Service worker: checks GitHub once a day and flags the toolbar icon when a newer version exists,
// so the notice is visible without opening the popup.
importScripts("version-check.js");

const ALARM_NAME = "update-check";
const DEFAULT_TITLE = "Teams Chat Exporter";

// The badge always mirrors the last known result (from this worker or from the popup's manual check)
async function applyBadge() {
  const { last } = await TceUpdate.getState();
  const current = chrome.runtime.getManifest().version;
  const hasUpdate = Boolean(last) && TceUpdate.compareVersions(last.version, current) > 0;

  await chrome.action.setBadgeText({ text: hasUpdate ? "!" : "" });
  if (hasUpdate) {
    await chrome.action.setBadgeBackgroundColor({ color: "#dc2626" });
    await chrome.action.setTitle({ title: `${DEFAULT_TITLE} — nova versão ${last.version} disponível` });
  } else {
    await chrome.action.setTitle({ title: DEFAULT_TITLE });
  }
}

async function checkNow() {
  try {
    const remoteVersion = await TceUpdate.fetchRemoteVersion();
    await TceUpdate.saveLast(remoteVersion);
  } catch {
    // offline or GitHub unreachable: keep the previous result and try again on the next alarm
  }
  await applyBadge();
}

async function ensureSchedule() {
  const { auto, last } = await TceUpdate.getState();
  if (!auto) {
    await chrome.alarms.clear(ALARM_NAME);
    await applyBadge();
    return;
  }
  if (!(await chrome.alarms.get(ALARM_NAME))) {
    await chrome.alarms.create(ALARM_NAME, { periodInMinutes: 24 * 60 });
  }
  if (TceUpdate.isStale(last)) {
    await checkNow();
  } else {
    await applyBadge();
  }
}

chrome.runtime.onInstalled.addListener(ensureSchedule);
chrome.runtime.onStartup.addListener(ensureSchedule);

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== ALARM_NAME) return;
  const { auto } = await TceUpdate.getState();
  if (auto) await checkNow();
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;
  if (changes[TceUpdate.KEY_AUTO]) {
    ensureSchedule();
  } else if (changes[TceUpdate.KEY_LAST]) {
    applyBadge();
  }
});
