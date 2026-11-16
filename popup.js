const statusBadge = document.getElementById("statusBadge");
const chatTitleEl = document.getElementById("chatTitle");
const msgCountEl = document.getElementById("msgCount");
const includeImagesEl = document.getElementById("includeImages");
const maxMessagesEl = document.getElementById("maxMessages");
const exportFormatEl = document.getElementById("exportFormat");
const startBtn = document.getElementById("startBtn");
const stopBtn = document.getElementById("stopBtn");
const downloadBtn = document.getElementById("downloadBtn");
const feedbackEl = document.getElementById("feedback");

// Tabs: "Exportar" (main) and "Sobre" (version, repository, updates)
const tabs = [
  { btn: document.getElementById("tabExport"), panel: document.getElementById("panelExport") },
  { btn: document.getElementById("tabAbout"), panel: document.getElementById("panelAbout") }
];
for (const tab of tabs) {
  tab.btn.addEventListener("click", () => {
    for (const t of tabs) {
      const active = t === tab;
      t.btn.classList.toggle("active", active);
      t.btn.setAttribute("aria-selected", String(active));
      t.panel.hidden = !active;
    }
  });
}

let pollTimer = null;
let currentTabId = null;
let collectStarted = false;

function setStatus(text, type) {
  statusBadge.textContent = text;
  statusBadge.className = `badge badge-${type}`;
}

function setFeedback(text, isError = false) {
  feedbackEl.textContent = text;
  feedbackEl.className = isError ? "feedback error" : "feedback";
}

function isTeamsUrl(url) {
  return (
    typeof url === "string" &&
    (url.startsWith("https://teams.microsoft.com/") ||
      url.startsWith("https://teams.live.com/") ||
      url.startsWith("https://teams.cloud.microsoft/"))
  );
}

async function ensureContentScript(tabId) {
  try {
    return await chrome.tabs.sendMessage(tabId, { type: "GET_INFO" });
  } catch {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
    return await chrome.tabs.sendMessage(tabId, { type: "GET_INFO" });
  }
}

async function refreshState() {
  if (!currentTabId) return null;
  try {
    const info = await chrome.tabs.sendMessage(currentTabId, { type: "GET_INFO" });
    chatTitleEl.textContent = info.chatTitle || "Chat sem título";
    msgCountEl.textContent = String(info.count || 0);

    if (info.running) {
      setStatus("Capturando...", "running");
      startBtn.disabled = true;
      stopBtn.disabled = false;
      downloadBtn.disabled = true;
      setFeedback("Rolando o histórico e capturando mensagens...");
    } else {
      startBtn.disabled = false;
      stopBtn.disabled = true;
      if ((info.count || 0) > 0) {
        setStatus("Concluído", "done");
        downloadBtn.disabled = false;
        const otherChat =
          info.collectedTitle && info.chatTitle && info.collectedTitle !== info.chatTitle;
        if (otherChat) {
          setFeedback(
            `As ${info.count} mensagens capturadas são de "${info.collectedTitle}". Clique em Capturar Chat para capturar o chat atual.`,
            true
          );
        } else {
          setFeedback(`Pronto para exportar ${info.count} mensagens.`);
        }
        collectStarted = false;
        stopPolling();
      } else {
        setStatus("Pronto", "idle");
        downloadBtn.disabled = true;
        if (collectStarted) {
          collectStarted = false;
          stopPolling();
          setFeedback("Nenhuma mensagem encontrada. Abra um chat no Teams e tente novamente.", true);
        }
      }
    }
    return info;
  } catch {
    stopPolling();
    setStatus("Erro", "error");
    setFeedback("Não foi possível comunicar com a aba do Teams.", true);
    return null;
  }
}

function startPolling() {
  stopPolling();
  pollTimer = setInterval(refreshState, 700);
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

function sanitizeFilename(name) {
  return (name || "teams-chat")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 80);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getAuthorColor(name) {
  const colors = [
    ["#e8ebfa", "#464775"],
    ["#e6f4ea", "#137333"],
    ["#fce8e6", "#c5221f"],
    ["#fef7e0", "#b06000"],
    ["#e4f7fb", "#007b83"],
    ["#f3e8fd", "#7627bb"],
    ["#fde7f3", "#b80672"],
    ["#e8f0fe", "#1967d2"]
  ];
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = (hash * 31 + name.charCodeAt(i)) >>> 0;
  }
  return colors[hash % colors.length];
}

function getInitials(name) {
  const parts = String(name || "?").trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function linkifyEscaped(escapedText) {
  return escapedText.replace(
    /(https?:\/\/[^\s<]+)/g,
    '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>'
  );
}

function formatDateLabel(dateObj) {
  return new Intl.DateTimeFormat("pt-BR", {
    day: "numeric",
    month: "long",
    year: "numeric"
  }).format(dateObj);
}

function formatTimeLabel(dateObj) {
  return new Intl.DateTimeFormat("pt-BR", {
    hour: "2-digit",
    minute: "2-digit"
  }).format(dateObj);
}

function formatMessages(chatTitle, messages, format, avatars = {}) {
  const exportedAt = new Date().toISOString();

  if (format === "html") {
    let lastDateKey = "";
    let lastAuthor = "";
    let lastTimeMs = 0;
    const chunks = [];

    for (const m of messages) {
      const parsedDate = m.timestamp ? new Date(m.timestamp) : null;
      const hasValidDate = parsedDate && !Number.isNaN(parsedDate.getTime());
      const dateKey = hasValidDate
        ? `${parsedDate.getFullYear()}-${parsedDate.getMonth()}-${parsedDate.getDate()}`
        : "";

      if (dateKey && dateKey !== lastDateKey) {
        chunks.push(`
      <div class="date-divider"><span>${escapeHtml(formatDateLabel(parsedDate))}</span></div>`);
        lastDateKey = dateKey;
        lastAuthor = "";
      }

      const isSystem = !m.author || m.author === "Desconhecido";
      const replyHtml = m.replyTo
        ? `<div class="quoted-reply">${
            m.replyTo.author || m.replyTo.timestamp
              ? `<div class="quoted-meta">${escapeHtml(m.replyTo.author || "")} ${escapeHtml(m.replyTo.timestamp || "")}</div>`
              : ""
          }<div class="quoted-text">${escapeHtml(m.replyTo.text || "").replace(/\n/g, "<br>")}</div></div>`
        : "";

      const bodyTextHtml = m.text ? linkifyEscaped(escapeHtml(m.text)).replace(/\n/g, "<br>") : "";
      const imagesHtml =
        !m.bodyHtml && Array.isArray(m.images) && m.images.length > 0
          ? `<div class="msg-images">${m.images
              .map(
                (src) =>
                  `<a href="${escapeHtml(src)}" target="_blank" rel="noopener noreferrer"><img class="chat-img" src="${escapeHtml(src)}" alt="Imagem do chat" loading="lazy" /></a>`
              )
              .join("")}</div>`
          : "";
      const attachmentsHtml =
        Array.isArray(m.attachments) && m.attachments.length > 0
          ? `<div class="attachments">${m.attachments
              .map((att) =>
                att.href
                  ? `<a class="attachment-card" href="${escapeHtml(att.href)}" target="_blank" rel="noopener noreferrer">&#128206; ${escapeHtml(att.title)}</a>`
                  : `<div class="attachment-card">&#128206; ${escapeHtml(att.title)}</div>`
              )
              .join("")}</div>`
          : "";

      const mainContentHtml = m.bodyHtml || `${bodyTextHtml}${imagesHtml}`;
      const bodyHtml = `${replyHtml}${mainContentHtml}${attachmentsHtml}`;

      const reactionsHtml =
        Array.isArray(m.reactions) && m.reactions.length > 0
          ? `<div class="reactions">${m.reactions
              .map((r) => `<span class="reaction-pill">${escapeHtml(r)}</span>`)
              .join("")}</div>`
          : "";

      if (isSystem) {
        chunks.push(`
      <div class="system-msg" data-search="${escapeHtml(m.text.toLowerCase())}">
        <span>${bodyHtml}</span>
      </div>`);
        lastAuthor = "";
        continue;
      }

      const timeMs = hasValidDate ? parsedDate.getTime() : 0;
      const isGrouped =
        !m.isUrgent &&
        m.author === lastAuthor &&
        (!hasValidDate || !lastTimeMs || timeMs - lastTimeMs <= 10 * 60 * 1000);

      const timeStr = hasValidDate ? formatTimeLabel(parsedDate) : escapeHtml(m.timestamp || "");
      const fullTs = hasValidDate ? parsedDate.toLocaleString("pt-BR") : m.timestamp || "";
      const [bg, fg] = getAuthorColor(m.author);
      const initials = escapeHtml(getInitials(m.author));
      const avatarUrl = avatars && avatars[m.author];
      const avatarHtml = avatarUrl
        ? `<div class="avatar"><img class="avatar-img" src="${escapeHtml(avatarUrl)}" alt="${initials}" /></div>`
        : `<div class="avatar" style="background:${bg};color:${fg};">${initials}</div>`;
      const searchAttr = escapeHtml(`${m.author} ${m.text}`.toLowerCase());
      const urgentBadge = m.isUrgent ? `<span class="urgent-badge">URGENT</span>` : "";

      if (isGrouped) {
        chunks.push(`
      <div class="msg-row grouped" data-search="${searchAttr}">
        <div class="avatar-spacer" title="${escapeHtml(fullTs)}">${escapeHtml(timeStr)}</div>
        <div class="msg-col">
          <div class="bubble">${bodyHtml}</div>
          ${reactionsHtml}
        </div>
      </div>`);
      } else {
        chunks.push(`
      <div class="msg-row" data-search="${searchAttr}">
        ${avatarHtml}
        <div class="msg-col">
          <div class="msg-meta">
            <span class="author">${escapeHtml(m.author)}</span>
            <span class="ts" title="${escapeHtml(fullTs)}">${escapeHtml(timeStr)}</span>
            ${urgentBadge}
          </div>
          <div class="bubble">${bodyHtml}</div>
          ${reactionsHtml}
        </div>
      </div>`);
      }

      lastAuthor = m.author;
      lastTimeMs = timeMs;
    }

    return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(chatTitle)} - Microsoft Teams</title>
  <style>
    :root {
      --bg: #1f1f1f;
      --topbar-bg: #141414;
      --border: #333333;
      --bubble-bg: #292929;
      --bubble-border: #333333;
      --text: #e8e8e8;
      --text-strong: #ffffff;
      --muted: #a6a6a6;
      --accent: #7f85f5;
      --system-bg: #292929;
      --input-bg: #292929;
      --divider: #383838;
    }
    body.light {
      --bg: #f5f5f5;
      --topbar-bg: #ffffff;
      --border: #e0e0e0;
      --bubble-bg: #ffffff;
      --bubble-border: #edebe9;
      --text: #242424;
      --text-strong: #242424;
      --muted: #616161;
      --accent: #5b5fc7;
      --system-bg: #ebebeb;
      --input-bg: #fafafa;
      --divider: #e0e0e0;
    }
    * { box-sizing: border-box; }
    body {
      font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
      background: var(--bg);
      color: var(--text);
      margin: 0;
      padding: 0;
    }
    .topbar {
      position: sticky;
      top: 0;
      z-index: 10;
      background: var(--topbar-bg);
      border-bottom: 1px solid var(--border);
      padding: 12px 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      flex-wrap: wrap;
    }
    .topbar-left {
      display: flex;
      align-items: center;
      gap: 12px;
      min-width: 0;
    }
    .topbar-right {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .chat-badge {
      width: 34px;
      height: 34px;
      border-radius: 6px;
      background: #5b5fc7;
      color: #fff;
      font-weight: 700;
      font-size: 14px;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }
    .title-wrap { min-width: 0; }
    .title-wrap h1 {
      margin: 0;
      font-size: 16px;
      font-weight: 600;
      color: var(--text-strong);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .subtitle {
      font-size: 12px;
      color: var(--muted);
    }
    .search-box {
      padding: 6px 12px;
      border: 1px solid var(--border);
      border-radius: 6px;
      font-size: 13px;
      width: 220px;
      max-width: 100%;
      background: var(--input-bg);
      color: var(--text);
    }
    .search-box:focus {
      outline: 2px solid var(--accent);
    }
    .theme-btn {
      padding: 6px 10px;
      border: 1px solid var(--border);
      border-radius: 6px;
      background: var(--input-bg);
      color: var(--text);
      font-size: 12px;
      cursor: pointer;
    }
    .chat-container {
      max-width: 920px;
      margin: 0 auto;
      padding: 16px 20px 40px;
      display: flex;
      flex-direction: column;
    }
    .date-divider {
      display: flex;
      align-items: center;
      margin: 18px 0 12px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 600;
    }
    .date-divider::before,
    .date-divider::after {
      content: "";
      flex: 1;
      height: 1px;
      background: var(--divider);
    }
    .date-divider span {
      padding: 0 12px;
    }
    .system-msg {
      text-align: center;
      font-size: 12px;
      color: var(--muted);
      margin: 8px auto;
      padding: 4px 12px;
      background: var(--system-bg);
      border-radius: 12px;
      max-width: 85%;
    }
    .msg-row {
      display: flex;
      align-items: flex-start;
      gap: 10px;
      margin-top: 12px;
    }
    .msg-row.grouped {
      margin-top: 3px;
    }
    .avatar {
      width: 32px;
      height: 32px;
      border-radius: 50%;
      font-size: 12px;
      font-weight: 600;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
      margin-top: 2px;
      overflow: hidden;
    }
    .avatar-img {
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
    }
    .avatar-spacer {
      width: 32px;
      flex-shrink: 0;
      font-size: 10px;
      color: transparent;
      text-align: right;
      padding-top: 9px;
      user-select: none;
    }
    .msg-row.grouped:hover .avatar-spacer {
      color: var(--muted);
    }
    .msg-col {
      min-width: 0;
      max-width: 760px;
    }
    .msg-meta {
      display: flex;
      align-items: baseline;
      gap: 8px;
      margin-bottom: 3px;
    }
    .author {
      font-size: 12px;
      font-weight: 600;
      color: var(--text-strong);
    }
    .ts {
      font-size: 11px;
      color: var(--muted);
    }
    .bubble {
      background: var(--bubble-bg);
      border: 1px solid var(--bubble-border);
      border-radius: 6px;
      padding: 8px 12px;
      font-size: 14px;
      line-height: 1.45;
      color: var(--text);
      word-break: break-word;
      display: inline-block;
    }
    .bubble a {
      color: var(--accent);
      text-decoration: none;
    }
    .bubble a:hover {
      text-decoration: underline;
    }
    .msg-images {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin-top: 6px;
    }
    .msg-images img,
    .chat-img {
      max-width: 100%;
      max-height: 420px;
      border-radius: 6px;
      border: 1px solid var(--border);
      object-fit: contain;
      display: block;
      margin-top: 6px;
    }
    .emoji-img {
      width: 22px;
      height: 22px;
      object-fit: cover;
      object-position: 0 0;
      vertical-align: -5px;
      display: inline-block;
      margin: 0 1px;
    }
    .attachments {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin-top: 8px;
    }
    .attachment-card {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 8px 12px;
      border: 1px solid var(--border);
      border-radius: 6px;
      background: rgba(128, 128, 128, 0.08);
      font-size: 13px;
      color: var(--text);
      text-decoration: none;
    }
    a.attachment-card:hover {
      border-color: var(--accent);
    }
    .mention {
      color: var(--accent);
      font-weight: 500;
    }
    .urgent-badge {
      color: #f87171;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.4px;
    }
    .quoted-reply {
      border: 1px solid var(--border);
      border-left: 3px solid var(--muted);
      background: rgba(128, 128, 128, 0.08);
      border-radius: 4px;
      padding: 6px 10px;
      margin-bottom: 8px;
      font-size: 13px;
    }
    .quoted-meta {
      font-size: 11px;
      color: var(--muted);
      margin-bottom: 3px;
    }
    .quoted-text {
      color: var(--text);
      opacity: 0.9;
    }
    .table-wrap {
      overflow-x: auto;
      margin: 6px 0;
      max-width: 100%;
    }
    .bubble table {
      border-collapse: collapse;
      font-size: 13px;
      width: 100%;
    }
    .bubble th,
    .bubble td {
      border: 1px solid var(--border);
      padding: 6px 10px;
      text-align: left;
      vertical-align: top;
    }
    .bubble th {
      background: rgba(128, 128, 128, 0.14);
      font-weight: 600;
    }
    .bubble p {
      margin: 0 0 6px 0;
    }
    .bubble p:last-child {
      margin-bottom: 0;
    }
    .reactions {
      display: flex;
      gap: 6px;
      margin-top: 4px;
      flex-wrap: wrap;
    }
    .reaction-pill {
      background: var(--bubble-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 2px 8px;
      font-size: 12px;
      color: var(--text);
    }
    .hidden {
      display: none !important;
    }
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <div class="chat-badge">${escapeHtml(getInitials(chatTitle))}</div>
      <div class="title-wrap">
        <h1>${escapeHtml(chatTitle)}</h1>
        <div class="subtitle">${messages.length} mensagens &bull; Exportado em ${escapeHtml(new Date(exportedAt).toLocaleString("pt-BR"))}</div>
      </div>
    </div>
    <div class="topbar-right">
      <input type="search" id="searchInput" class="search-box" placeholder="Buscar no chat..." />
      <button type="button" id="themeBtn" class="theme-btn">Tema claro</button>
    </div>
  </header>
  <main class="chat-container" id="chatContainer">
${chunks.join("\n")}
  </main>
  <script>
    const input = document.getElementById("searchInput");
    const rows = Array.from(document.querySelectorAll("[data-search]"));
    input.addEventListener("input", () => {
      const q = input.value.trim().toLowerCase();
      for (const el of rows) {
        el.classList.toggle("hidden", q !== "" && !el.dataset.search.includes(q));
      }
    });
    const themeBtn = document.getElementById("themeBtn");
    themeBtn.addEventListener("click", () => {
      const isLight = document.body.classList.toggle("light");
      themeBtn.textContent = isLight ? "Tema escuro" : "Tema claro";
    });
  </script>
</body>
</html>`;
  }

  if (format === "json") {
    return JSON.stringify(
      {
        chatTitle,
        exportedAt,
        totalMessages: messages.length,
        messages
      },
      null,
      2
    );
  }

  if (format === "md") {
    const lines = [
      `# ${chatTitle}`,
      "",
      `> Exportado em: ${exportedAt} | Total de mensagens: ${messages.length}`,
      ""
    ];
    for (const m of messages) {
      const ts = m.timestamp ? ` (${m.timestamp})` : "";
      lines.push(`### ${m.author}${ts}`);
      lines.push("");
      lines.push(m.text);
      lines.push("");
    }
    return lines.join("\n");
  }

  // Plain text
  const lines = [
    `Chat: ${chatTitle}`,
    `Exportado em: ${exportedAt}`,
    `Total de mensagens: ${messages.length}`,
    "=".repeat(50),
    ""
  ];
  for (const m of messages) {
    const ts = m.timestamp ? `[${m.timestamp}] ` : "";
    lines.push(`${ts}${m.author}:`);
    lines.push(m.text);
    lines.push("");
  }
  return lines.join("\n");
}

function triggerDownload(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

startBtn.addEventListener("click", async () => {
  if (!currentTabId) return;
  setStatus("Capturando...", "running");
  startBtn.disabled = true;
  stopBtn.disabled = false;
  downloadBtn.disabled = true;
  setFeedback("Iniciando captura...");

  try {
    await chrome.tabs.sendMessage(currentTabId, {
      type: "START_COLLECT",
      maxMessages: Number(maxMessagesEl.value) || 0,
      includeImages: includeImagesEl.checked
    });
    collectStarted = true;
    startPolling();
  } catch {
    setStatus("Erro", "error");
    startBtn.disabled = false;
    stopBtn.disabled = true;
    setFeedback("Não foi possível iniciar a captura. Recarregue a página do Teams.", true);
  }
});

stopBtn.addEventListener("click", async () => {
  if (!currentTabId) return;
  await chrome.tabs.sendMessage(currentTabId, { type: "STOP_COLLECT" });
  setFeedback("Parando captura...");
});

downloadBtn.addEventListener("click", async () => {
  if (!currentTabId) return;
  const format = exportFormatEl.value;
  downloadBtn.disabled = true;
  setFeedback("Preparando exportação...");

  try {
    const meta = await chrome.tabs.sendMessage(currentTabId, {
      type: "GET_EXPORT_META",
      maxMessages: Number(maxMessagesEl.value) || 0
    });

    if (!meta || !meta.total) {
      setFeedback("Nenhuma mensagem encontrada para exportar.", true);
      return;
    }

    // Messages come in chunks so large chats with embedded images fit in extension messaging
    const messages = [];
    while (messages.length < meta.total) {
      const chunk = await chrome.tabs.sendMessage(currentTabId, {
        type: "GET_MESSAGES_CHUNK",
        start: messages.length
      });
      if (!chunk || !Array.isArray(chunk.messages) || chunk.messages.length === 0) break;
      messages.push(...chunk.messages);
      setFeedback(`Preparando exportação... ${messages.length}/${meta.total}`);
    }

    if (messages.length < meta.total) {
      throw new Error(`apenas ${messages.length} de ${meta.total} mensagens foram recebidas`);
    }

    const content = formatMessages(meta.chatTitle, messages, format, meta.avatars || {});
    const mimeMap = {
      html: "text/html;charset=utf-8",
      md: "text/markdown;charset=utf-8",
      json: "application/json;charset=utf-8",
      txt: "text/plain;charset=utf-8"
    };

    const filename = `${sanitizeFilename(meta.chatTitle)}_${new Date()
      .toISOString()
      .slice(0, 10)}.${format}`;
    triggerDownload(content, filename, mimeMap[format]);
    setFeedback(`Arquivo ${filename} baixado com sucesso!`);
  } catch (err) {
    setFeedback(`Falha ao exportar: ${err && err.message ? err.message : err}`, true);
  } finally {
    downloadBtn.disabled = false;
  }
});

(async function init() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !isTeamsUrl(tab.url)) {
    setStatus("Inativo", "idle");
    chatTitleEl.textContent = "Abra o Teams Web";
    startBtn.disabled = true;
    setFeedback("Abra um chat em teams.microsoft.com, teams.live.com ou teams.cloud.microsoft.", true);
    return;
  }

  currentTabId = tab.id;
  try {
    await ensureContentScript(currentTabId);
    const info = await refreshState();
    if (info && info.running) {
      startPolling();
    }
  } catch {
    setStatus("Erro", "error");
    startBtn.disabled = true;
    setFeedback("Recarregue a página do Teams e tente novamente.", true);
  }
})();
