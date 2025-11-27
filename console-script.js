// Teams Chat Exporter - Console Script
// Paste this entire script into your browser console while on a Teams chat page
// Then call: exportTeamsChat() or exportTeamsChat("txt")

async function exportTeamsChat(format = "json") {
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  console.log("[Teams Chat Exporter] Starting export...");

  // Find a message to locate the scroll container
  const exampleContent = document.querySelector('div[id^="content-"]');
  if (!exampleContent) {
    console.error("[Teams Chat Exporter] No messages found. Make sure you're in a chat.");
    return;
  }

  // Find scroll container
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
      break;
    }
  }

  if (!scrollContainer) {
    console.error("[Teams Chat Exporter] Could not find scroll container.");
    return;
  }

  // Scroll to load history
  console.log("[Teams Chat Exporter] Scrolling to load chat history...");
  let lastHeight = -1;
  let stable = 0;
  const MAX_SCROLLS = 400;

  for (let i = 0; i < MAX_SCROLLS; i++) {
    scrollContainer.scrollTop = 0;
    scrollContainer.dispatchEvent(
      new WheelEvent("wheel", { deltaY: -1000, bubbles: true })
    );

    await sleep(700);

    const currentHeight = scrollContainer.scrollHeight;
    if (currentHeight === lastHeight) {
      stable++;
    } else {
      stable = 0;
    }
    lastHeight = currentHeight;

    if (stable >= 3) {
      console.log("[Teams Chat Exporter] Reached top of chat.");
      break;
    }

    if (i % 10 === 0) {
      console.log(`[Teams Chat Exporter] Still scrolling... (${i}/${MAX_SCROLLS})`);
    }
  }

  // Extract messages
  const messages = [];
  document.querySelectorAll('div[id^="content-"]').forEach(div => {
    const id = div.id;
    if (!id) return;

    const msgId = id.replace("content-", "");
    const text = div.innerText.trim();
    if (!text) return;

    const authorNode = document.getElementById("author-" + msgId);
    const timeNode = document.getElementById("timestamp-" + msgId);

    const sender = authorNode?.innerText?.trim() || "Unknown";
    const timestamp = timeNode
      ? {
          datetime: timeNode.getAttribute("datetime") || null,
          label: timeNode.getAttribute("title") || timeNode.getAttribute("aria-label") || timeNode.innerText?.trim() || null
        }
      : null;

    messages.push({ id: msgId, sender, timestamp, text });
  });

  // Sort chronologically
  messages.sort((a, b) => {
    const t1 = a.timestamp?.datetime ? new Date(a.timestamp.datetime) : 0;
    const t2 = b.timestamp?.datetime ? new Date(b.timestamp.datetime) : 0;
    return t1 - t2;
  });

  console.log(`[Teams Chat Exporter] Collected ${messages.length} messages.`);

  // Get chat name
  const chatNameEl = document.querySelector('[id^="chat-topic-"]');
  let chatName = chatNameEl?.innerText?.trim() || "teams_chat_export";
  chatName = chatName.replace(/[<>:"/\\|?*]+/g, "_").replace(/\s+/g, "_").slice(0, 80);

  const dateStamp = new Date().toISOString().slice(0, 10);

  // Generate content based on format
  let content, mimeType, fileExt;

  if (format === "txt") {
    content = messages.map(msg => {
      const time = msg.timestamp?.label || msg.timestamp?.datetime || "";
      return `[${time}] ${msg.sender}:\n${msg.text}\n`;
    }).join("\n");
    mimeType = "text/plain";
    fileExt = "txt";
  } else {
    content = JSON.stringify(messages, null, 2);
    mimeType = "application/json";
    fileExt = "json";
  }

  // Download
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${chatName}_${dateStamp}.${fileExt}`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  console.log(`[Teams Chat Exporter] Export complete: ${chatName}_${dateStamp}.${fileExt}`);
}

console.log("[Teams Chat Exporter] Script loaded. Run exportTeamsChat() for JSON or exportTeamsChat('txt') for text format.");
