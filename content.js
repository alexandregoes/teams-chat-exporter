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
    return document.title.replace(/\s*\|\s*Microsoft Teams.*$/i, "").trim() || "teams-chat";
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
      try {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 8000);
        const res = await fetch(src, { credentials: "include", signal: controller.signal });
        clearTimeout(timer);
        if (!res.ok) return src;
        const blob = await res.blob();
        return await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(typeof reader.result === "string" ? reader.result : src);
          reader.onerror = () => resolve(src);
          reader.readAsDataURL(blob);
        });
      } catch {
        return src;
      }
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

    return {
      author: qAuthor || "",
      timestamp: qTime || "",
      text: rawText
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
      const combined = emojiAlt && !txt.includes(emojiAlt) ? `${emojiAlt} ${txt}`.trim() : txt;
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
      for (const attr of Array.from(img.attributes)) {
        img.removeAttribute(attr.name);
      }
      img.setAttribute("src", dataUrl);
      img.setAttribute("alt", img.getAttribute("alt") || "Imagem do chat");
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
        const images =
          includeImages && msg.imageUrls.length > 0
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
          _batch: batchId,
          _within: withinIdx++
        });
        added++;
      }
    }
    return added;
  }

  function getOrderedMessages() {
    const list = Array.from(state.messagesMap.values());
    // Older batches (more negative _batch) come first; within each batch, top-to-bottom (_within)
    list.sort((a, b) => {
      if (a._batch !== b._batch) return a._batch - b._batch;
      return a._within - b._within;
    });
    return list.map(({ author, timestamp, text, bodyHtml, replyTo, reactions, attachments, isUrgent, images }) => ({
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

    const container = findMessageContainer();

    // First scroll all the way to the bottom so the newest messages are loaded
    if (autoScroll && container) {
      let bottomIdle = 0;
      while (!state.stopRequested && bottomIdle < 3) {
        const prevHeight = container.scrollHeight;
        container.scrollTop = container.scrollHeight;
        await sleep(400);
        const atBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 5;
        if (atBottom && container.scrollHeight === prevHeight) {
          bottomIdle++;
        } else {
          bottomIdle = 0;
        }
      }
    }

    await collectVisibleMessages(container, includeImages);

    if (autoScroll && container) {
      let idleRounds = 0;
      const maxIdleRounds = 6;

      while (!state.stopRequested && idleRounds < maxIdleRounds) {
        if (maxMessages > 0 && state.messagesMap.size >= maxMessages) {
          break;
        }

        const prevCount = state.messagesMap.size;
        const prevScrollTop = container.scrollTop;
        const prevScrollHeight = container.scrollHeight;

        // Scroll up toward the top of the message list
        const step = Math.max(300, Math.floor(container.clientHeight * 0.75));
        container.scrollTop = Math.max(0, container.scrollTop - step);

        await sleep(650);
        await collectVisibleMessages(container, includeImages);

        const grew =
          state.messagesMap.size > prevCount ||
          container.scrollHeight > prevScrollHeight ||
          container.scrollTop !== prevScrollTop;

        if (container.scrollTop === 0 && !grew) {
          idleRounds++;
        } else if (state.messagesMap.size === prevCount && container.scrollTop === 0) {
          idleRounds++;
        } else {
          idleRounds = 0;
        }
      }
    }

    state.running = false;
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === "GET_INFO") {
      sendResponse({
        chatTitle: getChatTitle(),
        running: state.running,
        count: state.messagesMap.size
      });
      return true;
    }

    if (msg.type === "START_COLLECT") {
      runCollection({
        autoScroll: Boolean(msg.autoScroll),
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

    if (msg.type === "GET_MESSAGES") {
      let messages = getOrderedMessages();
      const max = Number(msg.maxMessages) || 0;
      if (max > 0 && messages.length > max) {
        messages = messages.slice(-max);
      }
      sendResponse({
        chatTitle: state.chatTitle || getChatTitle(),
        messages,
        avatars: Object.fromEntries(state.avatarsMap)
      });
      return true;
    }
  });
})();
