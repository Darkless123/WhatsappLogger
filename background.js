chrome.runtime.onInstalled.addListener(() => {
  // Clear any old instances first to prevent duplicates
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: "logWhatsAppWithTrueMeta",
      title: "Send to Mandays Sheet",
      contexts: ["selection"]
    });
    console.log("Context menu item created successfully.");
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId === "logWhatsAppWithTrueMeta") {
    console.log("Right-click menu clicked! Text highlighted:", info.selectionText);

    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: extractWhatsAppMetadata,
    }, (results) => {
      // Catch script injection permission blocks
      if (chrome.runtime.lastError) {
        console.error("Chrome Script Injection Error:", chrome.runtime.lastError.message);
        return;
      }

      let meta = { sender: "WhatsApp Client", timestamp: "" };
      if (results && results[0] && results[0].result) {
        meta = results[0].result;
      }
      
      console.log("Extracted Meta from Page DOM:", meta);

      const payload = {
        text: info.selectionText,
        chat_sender: meta.sender,
        chat_datetime: meta.timestamp
      };

      // Dispatch payload to Flask
      fetch('http://localhost:5000/log-issue', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      })
      .then(response => response.json())
      .then(data => console.log('Flask Server Response:', data))
      .catch((error) => console.error('Network Error connecting to Flask:', error));
    });
  }
});

function extractWhatsAppMetadata() {
  let sender = "WhatsApp Client";
  let timestamp = "";

  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return { sender, timestamp };

  let element = selection.getRangeAt(0).startContainer;
  if (element.nodeType !== Node.ELEMENT_NODE) {
    element = element.parentNode;
  }

  const messageBlock = element.closest('[data-id], .focusable-list-item');
  if (!messageBlock) return { sender, timestamp };

  const copyableContainer = messageBlock.querySelector('[data-pre-plain-text]');
  if (copyableContainer) {
    const rawText = copyableContainer.getAttribute('data-pre-plain-text');
    if (rawText) {
      const match = rawText.match(/\[(.*?)\]\s*(.*?):/);
      if (match) {
        return { timestamp: match[1], sender: match[2].trim() };
      }
    }
  }

  const fallbackName = messageBlock.querySelector('span[dir="auto"]');
  if (fallbackName) {
    sender = fallbackName.innerText.trim();
  }

  return { sender, timestamp };
}