// Background service worker for SAP Hours Agent
console.log('[SAP Hours Agent] Service worker loaded');

// When extension icon is clicked, toggle the panel on the active tab
chrome.action.onClicked.addListener(async (tab) => {
  // Only toggle on SAP pages (content script handles the rest)
  try {
    await chrome.tabs.sendMessage(tab.id, { type: 'TOGGLE_PANEL' });
  } catch (e) {
    console.log('[SAP Hours Agent] Could not toggle panel:', e.message);
  }
});
