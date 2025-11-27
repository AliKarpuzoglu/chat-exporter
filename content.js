(() => {
  let mainButtonAdded = false;
  let isRunning = false;
  let exportFormat = localStorage.getItem("teams-chat-exporter-format") || "json";

  function log(...args) {
    console.log("[Teams Chat Exporter]", ...args);
  }

  function showSettingsModal() {
    // Remove existing modal if any
    const existing = document.getElementById("teams-chat-exporter-settings-modal");
    if (existing) {
      existing.remove();
      return;
    }

    const modal = document.createElement("div");
    modal.id = "teams-chat-exporter-settings-modal";
    Object.assign(modal.style, {
      position: "fixed",
      top: "0",
      left: "0",
      width: "100vw",
      height: "100vh",
      background: "rgba(0,0,0,0.7)",
      zIndex: "9999999",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    });

    const box = document.createElement("div");
    Object.assign(box.style, {
      background: "#1e1e1e",
      borderRadius: "8px",
      padding: "20px",
      width: "320px",
      boxShadow: "0 4px 20px rgba(0,0,0,0.3)"
    });

    const title = document.createElement("h2");
    title.textContent = "Settings";
    Object.assign(title.style, {
      margin: "0 0 20px 0",
      color: "#fff",
      fontSize: "18px"
    });

    // Export format section
    const formatLabel = document.createElement("div");
    formatLabel.textContent = "Export Format";
    Object.assign(formatLabel.style, {
      color: "#ccc",
      fontSize: "13px",
      marginBottom: "10px"
    });

    const formatContainer = document.createElement("div");
    Object.assign(formatContainer.style, {
      display: "flex",
      gap: "10px",
      marginBottom: "20px"
    });

    const createFormatBtn = (label, value) => {
      const btn = document.createElement("button");
      btn.textContent = label;
      btn.dataset.format = value;
      const isSelected = exportFormat === value;
      Object.assign(btn.style, {
        flex: "1",
        padding: "10px",
        border: isSelected ? "2px solid #2563EB" : "2px solid #444",
        background: isSelected ? "#2563EB" : "#333",
        color: "#fff",
        borderRadius: "6px",
        cursor: "pointer",
        fontWeight: "600",
        fontSize: "13px"
      });
      btn.addEventListener("click", () => {
        exportFormat = value;
        localStorage.setItem("teams-chat-exporter-format", value);
        formatContainer.querySelectorAll("button").forEach(b => {
          const sel = b.dataset.format === value;
          b.style.background = sel ? "#2563EB" : "#333";
          b.style.border = sel ? "2px solid #2563EB" : "2px solid #444";
        });
      });
      return btn;
    };

    formatContainer.appendChild(createFormatBtn("JSON", "json"));
    formatContainer.appendChild(createFormatBtn("TXT", "txt"));

    // Divider
    const divider = document.createElement("hr");
    Object.assign(divider.style, {
      border: "none",
      borderTop: "1px solid #444",
      margin: "15px 0"
    });

    // View Source button
    const sourceBtn = document.createElement("button");
    sourceBtn.textContent = "View Source Code";
    Object.assign(sourceBtn.style, {
      width: "100%",
      padding: "10px",
      background: "#6B7280",
      color: "#fff",
      border: "none",
      borderRadius: "6px",
      cursor: "pointer",
      fontWeight: "600",
      fontSize: "13px",
      marginBottom: "10px"
    });
    sourceBtn.addEventListener("click", () => {
      modal.remove();
      showSourceCode();
    });

    // Close button
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "Close";
    Object.assign(closeBtn.style, {
      width: "100%",
      padding: "10px",
      background: "#DC2626",
      color: "#fff",
      border: "none",
      borderRadius: "6px",
      cursor: "pointer",
      fontWeight: "600",
      fontSize: "13px"
    });
    closeBtn.addEventListener("click", () => modal.remove());

    box.appendChild(title);
    box.appendChild(formatLabel);
    box.appendChild(formatContainer);
    box.appendChild(divider);
    box.appendChild(sourceBtn);
    box.appendChild(closeBtn);
    modal.appendChild(box);

    modal.addEventListener("click", (e) => {
      if (e.target === modal) modal.remove();
    });

    document.body.appendChild(modal);
  }

  // Get a chat name from the header 
  function getChatName() {
    const el = document.querySelector('[id^="chat-topic-"]');
    if (!el) return "teams_chat_export";

    let name = el.innerText.trim();
    if (!name) return "teams_chat_export";

    // sanitize for filename: remove illegal characters, normalize spaces
    name = name.replace(/[<>:"/\\|?*]+/g, "_").replace(/\s+/g, "_");

    // avoid absurdly long filenames
    if (name.length > 80) {
      name = name.slice(0, 80);
    }

    return name || "teams_chat_export";
  }

  // Wait for messages to exist before injecting the main button
  function waitForMessagesAndInject() {
    const CHECK_INTERVAL_MS = 1500;

    const intervalId = setInterval(() => {
      const contentDiv = document.querySelector('div[id^="content-"]');
      if (contentDiv && !mainButtonAdded) {
        clearInterval(intervalId);
        injectMainButton();
      }
    }, CHECK_INTERVAL_MS);
  }

  function showSourceCode() {
    // Remove existing modal if any
    const existing = document.getElementById("teams-chat-exporter-source-modal");
    if (existing) {
      existing.remove();
      return;
    }

    const modal = document.createElement("div");
    modal.id = "teams-chat-exporter-source-modal";
    Object.assign(modal.style, {
      position: "fixed",
      top: "0",
      left: "0",
      width: "100vw",
      height: "100vh",
      background: "rgba(0,0,0,0.7)",
      zIndex: "9999999",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    });

    const box = document.createElement("div");
    Object.assign(box.style, {
      background: "#1e1e1e",
      borderRadius: "8px",
      padding: "20px",
      width: "80%",
      maxWidth: "900px",
      maxHeight: "80vh",
      display: "flex",
      flexDirection: "column",
      boxShadow: "0 4px 20px rgba(0,0,0,0.3)"
    });

    const header = document.createElement("div");
    Object.assign(header.style, {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: "15px"
    });

    const title = document.createElement("h2");
    title.textContent = "Source Code (content.js)";
    Object.assign(title.style, {
      margin: "0",
      color: "#fff",
      fontSize: "18px"
    });

    const closeBtn = document.createElement("button");
    closeBtn.textContent = "Close";
    Object.assign(closeBtn.style, {
      background: "#DC2626",
      color: "#fff",
      border: "none",
      padding: "8px 16px",
      borderRadius: "4px",
      cursor: "pointer",
      fontWeight: "600"
    });
    closeBtn.addEventListener("click", () => modal.remove());

    header.appendChild(title);
    header.appendChild(closeBtn);

    const info = document.createElement("p");
    info.innerHTML = `This is the complete source code running in this extension.<br><br>
      <strong>View on GitHub:</strong> <a href="https://github.com/AliKarpuzoglu/chat-exporter" target="_blank" style="color:#60a5fa;">github.com/AliKarpuzoglu/chat-exporter</a><br><br>
      You can also run this code directly in the browser console without installing the extension.`;
    Object.assign(info.style, {
      color: "#ccc",
      fontSize: "13px",
      marginBottom: "15px",
      lineHeight: "1.5"
    });

    const pre = document.createElement("pre");
    Object.assign(pre.style, {
      background: "#2d2d2d",
      padding: "15px",
      borderRadius: "4px",
      overflow: "auto",
      flex: "1",
      margin: "0",
      fontSize: "12px",
      lineHeight: "1.4",
      color: "#d4d4d4"
    });

    // Fetch the source code from the extension
    pre.textContent = "Loading source code...";

    fetch(chrome.runtime.getURL("content.js"))
      .then(r => r.text())
      .then(code => {
        pre.textContent = code;
      })
      .catch(() => {
        pre.textContent = "Could not load source. Visit github.com/AliKarpuzoglu/chat-exporter to view the code.";
      });

    box.appendChild(header);
    box.appendChild(info);
    box.appendChild(pre);
    modal.appendChild(box);

    modal.addEventListener("click", (e) => {
      if (e.target === modal) modal.remove();
    });

    document.body.appendChild(modal);
  }

  function injectMainButton() {
    if (mainButtonAdded) return;
    mainButtonAdded = true;

    // Settings button (gear icon)
    const settingsBtn = document.createElement("button");
    settingsBtn.innerHTML = "&#9881;"; // Gear icon
    settingsBtn.id = "teams-chat-exporter-settings-btn";
    settingsBtn.title = "Settings";
    Object.assign(settingsBtn.style, {
      position: "fixed",
      top: "10px",
      right: "175px",
      zIndex: "999999",
      padding: "8px 12px",
      fontSize: "18px",
      background: "#6B7280",
      color: "#fff",
      borderRadius: "6px",
      border: "none",
      cursor: "pointer",
      boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
      fontFamily: "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    });
    settingsBtn.addEventListener("click", showSettingsModal);
    document.body.appendChild(settingsBtn);

    const btn = document.createElement("button");
    btn.textContent = "Export Teams Chat";
    btn.id = "teams-chat-exporter-main-btn";

    Object.assign(btn.style, {
      position: "fixed",
      top: "10px",
      right: "10px",
      zIndex: "999999",
      padding: "10px 16px",
      fontSize: "13px",
      background: "#2563EB",
      color: "#fff",
      borderRadius: "6px",
      border: "none",
      cursor: "pointer",
      fontWeight: "600",
      boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
      fontFamily:
        "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    });

    btn.addEventListener("click", () => {
      if (isRunning) return;
      isRunning = true;
      btn.textContent = "Exporting…";
      btn.style.background = "#4B5563";
      startExportFlow().finally(() => {
        isRunning = false;
        btn.textContent = "Export Teams Chat";
        btn.style.background = "#2563EB";
      });
    });

    document.body.appendChild(btn);
    log("Main button injected.");
  }

  async function startExportFlow() {
    let stop = false;
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    // STOP button
    const stopBtn = document.createElement("button");
    stopBtn.textContent = "STOP SCROLL";
    stopBtn.id = "teams-chat-exporter-stop-btn";

    Object.assign(stopBtn.style, {
      position: "fixed",
      top: "50px",
      right: "10px",
      zIndex: "999999",
      padding: "8px 12px",
      fontSize: "12px",
      background: "#DC2626",
      color: "#fff",
      borderRadius: "6px",
      border: "none",
      cursor: "pointer",
      fontWeight: "600",
      boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
      fontFamily:
        "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    });

    stopBtn.addEventListener("click", () => {
      stop = true;
      stopBtn.textContent = "Stopping…";
      stopBtn.style.background = "#6B7280";
      log("Stop requested by user.");
    });

    document.body.appendChild(stopBtn);

    // 1) Find example message
    const exampleContent = document.querySelector('div[id^="content-"]');
    if (!exampleContent) {
      log("No messages found on screen. Aborting.");
      stopBtn.remove();
      return;
    }

    // 2) Find scroll container by walking up the DOM
    let scrollContainer = null;
    let el = exampleContent;
    while (el.parentElement) {
      el = el.parentElement;
      const style = getComputedStyle(el);
      if (
        el.scrollHeight > el.clientHeight + 20 &&
        ["auto", "scroll", "overlay"].includes(style.overflowY)
      ) {
        scrollContainer = el;
        scrollContainer.style.outline = "2px solid red"; // visual debug
        log("Scrolling container found:", scrollContainer);
        break;
      }
    }

    if (!scrollContainer) {
      log("Could not find a scrollable container. Aborting.");
      stopBtn.remove();
      return;
    }

    // 3) Scroll to load history
    log("Starting scroll to load chat history. Click STOP to cancel.");
    let lastHeight = -1;
    let stable = 0;
    const MAX_SCROLLS = 400;
    const SCROLL_PAUSE_MS = 700;

    for (let i = 0; i < MAX_SCROLLS && !stop; i++) {
      scrollContainer.scrollTop = 0;
      scrollContainer.dispatchEvent(
        new WheelEvent("wheel", { deltaY: -1000, bubbles: true })
      );

      await sleep(SCROLL_PAUSE_MS);

      const currentHeight = scrollContainer.scrollHeight;
      if (currentHeight === lastHeight) {
        stable++;
      } else {
        stable = 0;
      }
      lastHeight = currentHeight;

      if (stable >= 3) {
        log("Reached top or no more history loading.");
        break;
      }
    }

    log(
      "Scrolling finished (either reached top, max scrolls, or user stopped). Extracting messages…"
    );

    // 4) Extract messages
    const messages = [];
    const contentDivs = document.querySelectorAll('div[id^="content-"]');

    contentDivs.forEach(div => {
      const id = div.id; // e.g., "content-1750070214482"
      if (!id) return;

      const msgId = id.replace("content-", "");
      const text = div.innerText.trim();
      if (!text) return;

      const authorNode = document.getElementById("author-" + msgId);
      const timeNode = document.getElementById("timestamp-" + msgId);

      const sender =
        (authorNode && authorNode.innerText && authorNode.innerText.trim()) ||
        "Unknown";

      const timestamp = timeNode
        ? {
            datetime: timeNode.getAttribute("datetime") || null,
            label:
              timeNode.getAttribute("title") ||
              timeNode.getAttribute("aria-label") ||
              (timeNode.innerText && timeNode.innerText.trim()) ||
              null
          }
        : null;

      messages.push({ id: msgId, sender, timestamp, text });
    });

    // 5) Sort messages chronologically by timestamp before export
    messages.sort((a, b) => {
      const t1 = a.timestamp?.datetime ? new Date(a.timestamp.datetime) : 0;
      const t2 = b.timestamp?.datetime ? new Date(b.timestamp.datetime) : 0;
      return t1 - t2;
    });

    // Re-read format from localStorage in case it changed
    const currentFormat = localStorage.getItem("teams-chat-exporter-format") || "json";
    log(`Collected ${messages.length} messages. Exporting as ${currentFormat.toUpperCase()}…`);

    // 6) Download file with chat name + date in filename
    try {
      const chatName = getChatName();
      const dateStamp = new Date().toISOString().slice(0, 10); // YYYY-MM-DD

      let content, mimeType, fileExt;

      if (currentFormat === "txt") {
        // Format as readable text
        content = messages.map(msg => {
          const time = msg.timestamp?.label || msg.timestamp?.datetime || "";
          return `[${time}] ${msg.sender}:\n${msg.text}\n`;
        }).join("\n");
        mimeType = "text/plain";
        fileExt = "txt";
      } else {
        // Default to JSON
        content = JSON.stringify(messages, null, 2);
        mimeType = "application/json";
        fileExt = "json";
      }

      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const fileName = `${chatName}_${dateStamp}.${fileExt}`;

      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      log("Export complete:", fileName);
    } catch (e) {
      console.error("[Teams Chat Exporter] Error during export:", e);
    }

    stopBtn.remove();
  }

  // Start when content script loads
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", waitForMessagesAndInject);
  } else {
    waitForMessagesAndInject();
  }
})();
