(() => {
  if (window.__teamsChatExporterLoaded) return;
  window.__teamsChatExporterLoaded = true;

  const state = {
    running: false,
    stopRequested: false,
    chatTitle: "",
    messagesMap: new Map(),
    avatarsMap: new Map(),
    imageCache: new Map(),
    exportList: null,
    batchCounter: 0
  };

  function getChatTitle() {
    const selectors = [
      '[data-tid="chat-title"]',
      '[data-tid="entity-header-title"]',
      '[data-tid="thread-header-title"]',
      'h2[data-tid*="title"]',
      'header h2'
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el && el.textContent && el.textContent.trim()) {
        return el.textContent.trim();
      }
    }
    return (
      document.title
        .replace(/^\(\d+\+?\)\s*/, "")
        .replace(/\s*\|\s*Microsoft Teams.*$/i, "")
        .trim() || "teams-chat"
    );
  }

  function findMessageContainer() {
    const containerSelectors = [
      '[data-tid="message-pane-list-viewport"]',
      '[data-tid="chat-pane-list"]',
      '[data-tid="message-pane-list"]',
      'div[role="list"][aria-label*="message" i]',
      'div[role="list"]'
    ];

    for (const sel of containerSelectors) {
      const el = document.querySelector(sel);
      if (el) {
        return findScrollableParent(el) || el;
      }
    }

    const firstMsg = document.querySelector('[data-tid="chat-pane-item"], [data-tid="message-body"]');
    if (firstMsg) {
      return findScrollableParent(firstMsg);
    }

    return null;
  }

  function findScrollableParent(el) {
    let cur = el;
    while (cur && cur !== document.body) {
      const style = window.getComputedStyle(cur);
      const overflowY = style.overflowY;
      if ((overflowY === "auto" || overflowY === "scroll") && cur.scrollHeight > cur.clientHeight + 10) {
        return cur;
      }
      cur = cur.parentElement;
    }
    return null;
  }

  function getMessageNodes(root) {
    const scope = root || document;
    const selectors = [
      '[data-tid="chat-pane-item"]',
      '[data-tid="message-pane-item"]',
      'div[role="listitem"]'
    ];

    for (const sel of selectors) {
      const nodes = Array.from(scope.querySelectorAll(sel));
      if (nodes.length > 0) {
        return nodes;
      }
    }

    return Array.from(scope.querySelectorAll('[data-tid="message-body"]')).map(
      (body) => body.closest('[role="listitem"]') || body.parentElement
    );
  }

  function extractText(el) {
    if (!el) return "";
    return el.innerText.replace(/\u200B/g, "").trim();
  }

  function extractImageUrls(node, bodyEl) {
    const scope = bodyEl || node;
    const imgs = Array.from(scope.querySelectorAll("img"));
    const urls = [];
    const seen = new Set();

    for (const img of imgs) {
      const src = img.currentSrc || img.src;
      if (!src || src.startsWith("data:image/svg+xml")) continue;
      if (img.closest('[data-tid*="emoji"], [data-tid*="reaction"], [data-tid*="avatar"]')) continue;
      if (img.hasAttribute("data-emoji") || img.className?.toString().includes("emoji")) continue;

      const w = img.naturalWidth || img.width || 0;
      const h = img.naturalHeight || img.height || 0;
      if (w > 0 && h > 0 && w <= 40 && h <= 40) continue;

      if (!seen.has(src)) {
        seen.add(src);
        urls.push(src);
      }
    }
    return urls;
  }

  async function urlToDataUrl(src) {
    if (src.startsWith("data:")) return src;
    if (state.imageCache.has(src)) return state.imageCache.get(src);

    const task = (async () => {
      // Authenticated first (chat images); public CDNs (emoticons, icons) reject
      // credentialed cross-origin requests, so retry anonymously.
      for (const credentials of ["include", "omit"]) {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 8000);
        try {
          const res = await fetch(src, { credentials, signal: controller.signal });
          if (!res.ok) continue;
          const blob = await res.blob();
          const dataUrl = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(typeof reader.result === "string" ? reader.result : null);
            reader.onerror = () => resolve(null);
            reader.readAsDataURL(blob);
          });
          if (dataUrl) return dataUrl;
        } catch {
          // try next mode
        } finally {
          clearTimeout(timer);
        }
      }
      return src;
    })();

    state.imageCache.set(src, task);
    return task;
  }

  const ALLOWED_TAGS = new Set([
    "P", "BR", "STRONG", "B", "EM", "I", "U", "S", "CODE", "PRE",
    "UL", "OL", "LI", "BLOCKQUOTE", "TABLE", "THEAD", "TBODY", "TFOOT",
    "TR", "TH", "TD", "A", "IMG", "DIV", "SPAN", "H1", "H2", "H3", "H4",
    "HR", "SUB", "SUP", "MARK", "DL", "DT", "DD"
  ]);

  function extractQuotedReply(node, bodyEl) {
    const scope = node || bodyEl;
    if (!scope) return null;
    const quoted = scope.querySelector(
      '[data-tid="quoted-reply-card"], [data-tid="reply-card"], [data-tid*="quoted-reply"]'
    );
    if (!quoted) return null;

    const qAuthor = extractText(
      quoted.querySelector('[data-tid*="author"], [class*="author"]')
    );
    const qTimeEl = quoted.querySelector("time, [data-tid*='timestamp']");
    const qTime = qTimeEl
      ? qTimeEl.getAttribute("datetime") || extractText(qTimeEl)
      : "";
    const rawText = extractText(quoted);
    if (!rawText) return null;

    // The card text usually starts with "Author\nTimestamp\n..."; lift those
    // lines into the metadata instead of repeating them in the quoted text.
    let author = qAuthor || "";
    const lines = rawText.split("\n").map((l) => l.trim()).filter(Boolean);
    if (!author && qTime && lines.length >= 3 && lines[1] === qTime) {
      author = lines[0];
    }
    while (lines.length > 1 && (lines[0] === author || lines[0] === qTime)) {
      lines.shift();
    }

    return {
      author,
      timestamp: qTime || "",
      text: lines.join("\n")
    };
  }

  function extractReactions(node) {
    const pills = Array.from(
      node.querySelectorAll('[data-tid*="reaction-pill"], [data-tid*="reaction-button"], [data-tid*="reaction-count"]')
    );
    const out = [];
    const seen = new Set();
    for (const pill of pills) {
      const emojiImg = pill.querySelector("img");
      const emojiAlt = emojiImg
        ? emojiImg.getAttribute("alt") || emojiImg.getAttribute("aria-label") || ""
        : "";
      const txt = extractText(pill).replace(/\s+/g, " ");
      // Teams exposes "1 Like reaction." / "2 Like reactions. 2"; keep just emoji + count
      const countMatch = txt.match(/\d+/);
      let combined;
      if (emojiAlt && countMatch) {
        combined = `${emojiAlt} ${countMatch[0]}`;
      } else {
        combined = emojiAlt && !txt.includes(emojiAlt) ? `${emojiAlt} ${txt}`.trim() : txt;
      }
      if (combined && !seen.has(combined)) {
        seen.add(combined);
        out.push(combined);
      }
    }
    return out;
  }

  function extractAttachments(node, bodyEl) {
    const candidates = Array.from(
      node.querySelectorAll(
        '[data-tid*="attachment"], [data-tid*="file-card"], [data-tid*="file-preview"]'
      )
    );
    const out = [];
    const seen = new Set();
    for (const el of candidates) {
      if (bodyEl && bodyEl.contains(el)) continue;
      const link = el.querySelector("a[href]") || (el.tagName === "A" ? el : null);
      let href = link ? link.getAttribute("href") || "" : "";
      if (href) {
        try {
          href = new URL(href, window.location.origin).href;
        } catch {
          // keep as-is
        }
      }
      const title = extractText(el).replace(/\s+/g, " ").slice(0, 160);
      if (!title && !href) continue;
      const key = `${title}|${href}`;
      if (!seen.has(key)) {
        seen.add(key);
        out.push({ title: title || href, href });
      }
    }
    return out;
  }

  async function buildCleanBodyHtml(bodyEl, includeImages) {
    if (!bodyEl) return "";
    const clone = bodyEl.cloneNode(true);

    // Mark elements that are inline or inside a horizontal flex row in the live DOM
    const origEls = Array.from(bodyEl.querySelectorAll("*"));
    const cloneEls = Array.from(clone.querySelectorAll("*"));
    for (let i = 0; i < origEls.length && i < cloneEls.length; i++) {
      const o = origEls[i];
      const c = cloneEls[i];
      const tag = o.tagName;
      if (
        o.hasAttribute("itemprop") ||
        String(o.getAttribute("itemtype") || "").toLowerCase().includes("mention") ||
        String(o.getAttribute("data-tid") || "").toLowerCase().includes("mention") ||
        String(o.className || "").toLowerCase().includes("mention")
      ) {
        c.setAttribute("data-keep-mention", "1");
      }
      if (tag === "DIV" || tag === "P") {
        const cs = window.getComputedStyle(o);
        const pcs = o.parentElement ? window.getComputedStyle(o.parentElement) : null;
        const isInline = cs.display.startsWith("inline");
        const isFlexRow =
          pcs &&
          (pcs.display === "flex" || pcs.display === "inline-flex") &&
          pcs.flexDirection.startsWith("row");
        if (isInline || isFlexRow) {
          c.setAttribute("data-inline", "1");
        }
      }
    }

    // Remove UI-only and quoted-reply elements (quoted reply is rendered separately)
    const removeSelectors = [
      "script", "style", "iframe", "object", "embed", "svg", "button",
      '[data-tid="quoted-reply-card"]',
      '[data-tid="reply-card"]',
      '[data-tid*="quoted-reply"]',
      '[data-tid*="translate"]',
      '[data-tid*="reaction"]'
    ];
    for (const sel of removeSelectors) {
      for (const el of Array.from(clone.querySelectorAll(sel))) {
        el.remove();
      }
    }

    // Process images (emojis -> Unicode alt or inline emoji-img; content images -> Base64)
    const imgs = Array.from(clone.querySelectorAll("img"));
    for (const img of imgs) {
      const src = img.currentSrc || img.getAttribute("src") || "";
      const nw = img.naturalWidth || img.width || 0;
      const nh = img.naturalHeight || img.height || 0;
      const isSpriteEmoji = nw > 0 && nw <= 64 && nh >= nw * 3;
      const isEmoji =
        isSpriteEmoji ||
        img.hasAttribute("data-emoji") ||
        String(img.getAttribute("itemtype") || "").toLowerCase().includes("emoji") ||
        Boolean(img.closest('[data-tid*="emoji"], [data-tid*="emoticon"]')) ||
        String(img.className || "").toLowerCase().includes("emoji");

      if (isEmoji) {
        const alt = (
          img.getAttribute("alt") ||
          img.getAttribute("aria-label") ||
          img.getAttribute("title") ||
          ""
        ).trim();
        const hasUnicodeEmoji = /[^\x00-\x7F]/.test(alt);
        if (hasUnicodeEmoji && !isSpriteEmoji) {
          img.replaceWith(document.createTextNode(alt));
          continue;
        }
        if (src && includeImages && !src.startsWith("data:image/svg+xml")) {
          const dataUrl = await urlToDataUrl(src);
          for (const attr of Array.from(img.attributes)) img.removeAttribute(attr.name);
          img.setAttribute("src", dataUrl);
          img.setAttribute("alt", alt || "emoji");
          img.setAttribute("class", "emoji-img");
          continue;
        }
        img.replaceWith(document.createTextNode(alt));
        continue;
      }

      if (!src || src.startsWith("data:image/svg+xml") || img.closest('[data-tid*="avatar"]')) {
        img.remove();
        continue;
      }

      if (!includeImages) {
        img.remove();
        continue;
      }

      const dataUrl = await urlToDataUrl(src);
      const originalAlt = (img.getAttribute("alt") || "").trim();
      for (const attr of Array.from(img.attributes)) {
        img.removeAttribute(attr.name);
      }
      img.setAttribute("src", dataUrl);
      img.setAttribute("alt", originalAlt || "Imagem do chat");
      img.setAttribute("class", "chat-img");
      img.setAttribute("loading", "lazy");
    }

    // Collapse @mentions into clean inline <span class="mention">Full Name</span>
    const mentionSelectors = '[data-keep-mention="1"], [itemtype*="mention" i], [itemprop="mri" i], [data-tid*="mention" i], [class*="mention" i]';
    for (const mEl of Array.from(clone.querySelectorAll(mentionSelectors))) {
      if (!clone.contains(mEl)) continue;
      // Skip if nested inside another mention already handled
      const mentionText = (mEl.textContent || "").replace(/\u200B/g, "").replace(/\s+/g, " ").trim();
      if (!mentionText) continue;
      const span = document.createElement("span");
      span.setAttribute("class", "mention");
      span.setAttribute("data-keep-mention", "1");
      span.textContent = mentionText;
      mEl.replaceWith(span);
    }

    // Convert live-inline DIV/P elements into SPAN so they don't break lines
    for (const inl of Array.from(clone.querySelectorAll('[data-inline="1"]'))) {
      if (!clone.contains(inl)) continue;
      const span = document.createElement("span");
      span.append(...Array.from(inl.childNodes));
      if (inl.getAttribute("data-keep-mention") === "1") {
        span.setAttribute("data-keep-mention", "1");
      }
      inl.replaceWith(span);
    }

    // Sanitize remaining elements and attributes
    const allEls = Array.from(clone.querySelectorAll("*"));
    for (const el of allEls) {
      if (!clone.contains(el)) continue;
      const tag = el.tagName;
      if (!ALLOWED_TAGS.has(tag)) {
        el.replaceWith(...Array.from(el.childNodes));
        continue;
      }

      if (tag === "IMG") continue;

      const isMention = el.getAttribute("data-keep-mention") === "1";

      const href = tag === "A" ? el.getAttribute("href") : null;
      const colspan = (tag === "TD" || tag === "TH") ? el.getAttribute("colspan") : null;
      const rowspan = (tag === "TD" || tag === "TH") ? el.getAttribute("rowspan") : null;

      for (const attr of Array.from(el.attributes)) {
        el.removeAttribute(attr.name);
      }

      if (tag === "A" && href) {
        let resolvedHref = href;
        try {
          resolvedHref = new URL(href, window.location.origin).href;
        } catch {
          // keep as-is
        }
        if (/^(https?:|mailto:|tel:)/i.test(resolvedHref)) {
          el.setAttribute("href", resolvedHref);
          el.setAttribute("target", "_blank");
          el.setAttribute("rel", "noopener noreferrer");
        }
      }
      if (isMention) {
        el.setAttribute("class", "mention");
      }
      if (colspan) el.setAttribute("colspan", colspan);
      if (rowspan) el.setAttribute("rowspan", rowspan);
    }

    // Wrap tables in scrollable container
    for (const tbl of Array.from(clone.querySelectorAll("table"))) {
      const wrap = document.createElement("div");
      wrap.setAttribute("class", "table-wrap");
      tbl.parentNode.insertBefore(wrap, tbl);
      wrap.appendChild(tbl);
    }

    return clone.innerHTML.replace(/\u200B/g, "").trim();
  }

  function parseMessageNode(node) {
    const bodyEl =
      node.querySelector('[data-tid="message-body"]') ||
      node.querySelector('[id^="content-"]') ||
      node.querySelector('[data-tid="message-content"]');

    const text = extractText(bodyEl);
    const imageUrls = extractImageUrls(node, bodyEl);
    if (!text && imageUrls.length === 0) return null;

    const authorEl =
      node.querySelector('[data-tid="message-author-name"]') ||
      node.querySelector('[data-tid="post-author"]') ||
      node.querySelector('[data-tid="thread-author-name"]');
    const author = extractText(authorEl) || "Desconhecido";

    const timeEl =
      node.querySelector("time[datetime]") ||
      node.querySelector('[data-tid="message-timestamp"]') ||
      node.querySelector("time");
    const timestamp = timeEl
      ? timeEl.getAttribute("datetime") || extractText(timeEl)
      : "";

    const rawId =
      node.getAttribute("data-mid") ||
      node.getAttribute("id") ||
      node.querySelector("[data-mid]")?.getAttribute("data-mid") ||
      `${author}|${timestamp}|${(text || imageUrls[0] || "").slice(0, 80)}`;

    const rect = node.getBoundingClientRect();
    const replyTo = extractQuotedReply(node, bodyEl);
    const reactions = extractReactions(node);
    const attachments = extractAttachments(node, bodyEl);
    const isUrgent = Boolean(
      node.querySelector('[data-tid*="urgent"], [data-tid*="importance"]')
    );

    const avatarImg =
      node.querySelector('[data-tid*="avatar"] img') ||
      Array.from(node.querySelectorAll("img")).find((im) => {
        if (bodyEl && bodyEl.contains(im)) return false;
        const w = im.naturalWidth || im.width || 0;
        const h = im.naturalHeight || im.height || 0;
        return w >= 20 && w <= 80 && h >= 20 && h <= 80;
      });
    const avatarSrc = avatarImg ? avatarImg.currentSrc || avatarImg.src || "" : "";

    return {
      id: rawId,
      author,
      timestamp,
      text,
      bodyEl,
      replyTo,
      reactions,
      attachments,
      isUrgent,
      avatarSrc,
      imageUrls,
      top: rect.top
    };
  }

  async function collectVisibleMessages(container, includeImages) {
    const nodes = getMessageNodes(container || document);
    const parsed = [];
    for (const node of nodes) {
      const msg = parseMessageNode(node);
      if (msg) parsed.push(msg);
    }

    // Sort visible batch top-to-bottom
    parsed.sort((a, b) => a.top - b.top);

    const batchId = state.batchCounter--;
    let added = 0;
    let withinIdx = 0;
    for (const msg of parsed) {
      if (
        includeImages &&
        msg.avatarSrc &&
        msg.author !== "Desconhecido" &&
        !state.avatarsMap.has(msg.author) &&
        !msg.avatarSrc.startsWith("data:image/svg+xml")
      ) {
        const avatarDataUrl = await urlToDataUrl(msg.avatarSrc);
        if (avatarDataUrl && avatarDataUrl.startsWith("data:")) {
          state.avatarsMap.set(msg.author, avatarDataUrl);
        }
      }

      if (!state.messagesMap.has(msg.id)) {
        const bodyHtml = await buildCleanBodyHtml(msg.bodyEl, includeImages);
        // Images already embedded in bodyHtml are not stored a second time
        const images =
          includeImages && !bodyHtml && msg.imageUrls.length > 0
            ? await Promise.all(msg.imageUrls.map((u) => urlToDataUrl(u)))
            : [];

        state.messagesMap.set(msg.id, {
          id: msg.id,
          author: msg.author,
          timestamp: msg.timestamp,
          text: msg.text,
          bodyHtml,
          replyTo: msg.replyTo,
          reactions: msg.reactions,
          attachments: msg.attachments,
          isUrgent: msg.isUrgent,
          images,
          _key: getSortKey(msg.id, msg.timestamp),
          _batch: batchId,
          _within: withinIdx++
        });
        added++;
      }
    }
    return added;
  }

  // Chronological key: ISO timestamp when available, else a 13-digit epoch message id
  function getSortKey(id, timestamp) {
    if (timestamp && /^\d{4}-\d{2}-\d{2}T/.test(timestamp)) {
      const t = Date.parse(timestamp);
      if (!Number.isNaN(t)) return t;
    }
    if (/^\d{13}$/.test(String(id))) return Number(id);
    return null;
  }

  function getOrderedMessages() {
    const list = Array.from(state.messagesMap.values());
    // Baseline: older batches (more negative _batch) first; within each batch, top-to-bottom
    list.sort((a, b) => {
      if (a._batch !== b._batch) return a._batch - b._batch;
      return a._within - b._within;
    });

    // Messages without a time key inherit the key of their nearest neighbour,
    // so they stay next to the messages they were seen with.
    const keys = list.map((m) => m._key);
    let carry = null;
    for (let i = 0; i < keys.length; i++) {
      if (keys[i] != null) carry = keys[i];
      else keys[i] = carry;
    }
    carry = null;
    for (let i = keys.length - 1; i >= 0; i--) {
      if (list[i]._key != null) carry = list[i]._key;
      if (keys[i] == null) keys[i] = carry;
    }

    const order = list.map((_, i) => i);
    order.sort((a, b) => (keys[a] ?? 0) - (keys[b] ?? 0) || a - b);

    return order.map((i) => list[i]).map(({ author, timestamp, text, bodyHtml, replyTo, reactions, attachments, isUrgent, images }) => ({
      author,
      timestamp,
      text,
      bodyHtml: bodyHtml || "",
      replyTo: replyTo || null,
      reactions: reactions || [],
      attachments: attachments || [],
      isUrgent: Boolean(isUrgent),
      images: images || []
    }));
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async function runCollection({ autoScroll, maxMessages, includeImages }) {
    if (state.running) return;
    state.running = true;
    state.stopRequested = false;
    state.messagesMap.clear();
    state.avatarsMap.clear();
    state.imageCache.clear();
    state.batchCounter = 0;
    state.chatTitle = getChatTitle();

    state.exportList = null;

    try {
      const container = findMessageContainer();

      // First scroll all the way to the bottom so the newest messages are loaded.
      // "Settled" = position and height stop changing (works for reversed scrollers too).
      if (autoScroll && container) {
        let settled = 0;
        for (let i = 0; i < 60 && !state.stopRequested && settled < 3; i++) {
          const prevHeight = container.scrollHeight;
          container.scrollTop = container.scrollHeight;
          const afterSet = container.scrollTop;
          await sleep(400);
          if (
            container.scrollHeight === prevHeight &&
            Math.abs(container.scrollTop - afterSet) < 2
          ) {
            settled++;
          } else {
            settled = 0;
          }
        }
      }

      await collectVisibleMessages(container, includeImages);

      if (autoScroll && container) {
        // Stop after ~8s with no movement, no growth and no new messages
        // (top reached, or the scroller is stuck).
        const maxIdleRounds = 12;
        // Safety net: stop if the list keeps moving but yields nothing new for ~26s.
        const maxRoundsWithoutNew = 40;
        let idleRounds = 0;
        let roundsWithoutNew = 0;

        while (
          !state.stopRequested &&
          idleRounds < maxIdleRounds &&
          roundsWithoutNew < maxRoundsWithoutNew
        ) {
          if (maxMessages > 0 && state.messagesMap.size >= maxMessages) {
            break;
          }

          const prevCount = state.messagesMap.size;
          const prevScrollTop = container.scrollTop;
          const prevScrollHeight = container.scrollHeight;

          // Scroll up toward the top of the message list (the browser clamps the value)
          const step = Math.max(300, Math.floor(container.clientHeight * 0.75));
          container.scrollTop = prevScrollTop - step;

          await sleep(650);
          await collectVisibleMessages(container, includeImages);

          const gotNew = state.messagesMap.size > prevCount;
          const moved =
            container.scrollHeight !== prevScrollHeight ||
            Math.abs(container.scrollTop - prevScrollTop) > 1;

          roundsWithoutNew = gotNew ? 0 : roundsWithoutNew + 1;
          idleRounds = gotNew || moved ? 0 : idleRounds + 1;
        }
      }
    } finally {
      state.running = false;
    }
  }

  // Splits the export into chunks so no single extension message gets near Chrome's 64 MB limit
  const CHUNK_CHAR_BUDGET = 8 * 1024 * 1024;

  function approxMessageSize(m) {
    let size = 500 + m.text.length + m.bodyHtml.length;
    for (const img of m.images) size += img.length;
    return size;
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === "GET_INFO") {
      sendResponse({
        chatTitle: getChatTitle(),
        collectedTitle: state.chatTitle,
        running: state.running,
        count: state.messagesMap.size
      });
      return true;
    }

    if (msg.type === "START_COLLECT") {
      runCollection({
        // The whole chat is always collected: scroll to the bottom, then up through the history
        autoScroll: true,
        maxMessages: Number(msg.maxMessages) || 0,
        includeImages: msg.includeImages !== false
      }).then(() => {
        // Done
      });
      sendResponse({ started: true });
      return true;
    }

    if (msg.type === "STOP_COLLECT") {
      state.stopRequested = true;
      sendResponse({ stopping: true });
      return true;
    }

    if (msg.type === "GET_EXPORT_META") {
      let messages = getOrderedMessages();
      const max = Number(msg.maxMessages) || 0;
      if (max > 0 && messages.length > max) {
        messages = messages.slice(-max);
      }
      state.exportList = messages;
      sendResponse({
        chatTitle: state.chatTitle || getChatTitle(),
        total: messages.length,
        avatars: Object.fromEntries(state.avatarsMap)
      });
      return true;
    }

    if (msg.type === "GET_MESSAGES_CHUNK") {
      const list = state.exportList || [];
      const start = Math.max(0, Number(msg.start) || 0);
      const chunk = [];
      let size = 0;
      for (let i = start; i < list.length; i++) {
        const s = approxMessageSize(list[i]);
        if (chunk.length > 0 && size + s > CHUNK_CHAR_BUDGET) break;
        chunk.push(list[i]);
        size += s;
      }
      if (start + chunk.length >= list.length) {
        state.exportList = null;
      }
      sendResponse({ messages: chunk });
      return true;
    }
  });
})();
