// Background service worker for SAP Hours Agent
console.log('[SAP Hours Agent] Service worker loaded');

// When extension icon is clicked, toggle the panel on the active tab
chrome.action.onClicked.addListener(async (tab) => {
  try {
    await chrome.tabs.sendMessage(tab.id, { type: 'TOGGLE_PANEL' });
  } catch (e) {
    console.log('[SAP Hours Agent] Could not toggle panel:', e.message);
  }
});

// After Easy Auth login, /api/token redirects to /api/token/done#t=<token>&exp=<expiry>
// We read token from the fragment in tab.url, store it, and close the tab.
chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  if (changeInfo.status !== 'complete') return;
  if (!tab.url) return;

  let url;
  try { url = new URL(tab.url); } catch { return; }

  if (url.hostname !== 'sap-hours-proxy.azurewebsites.net') return;
  if (!url.pathname.endsWith('/api/token/done')) return;

  const params = new URLSearchParams(url.hash.substring(1));
  const token = params.get('t');
  const expiry = parseInt(params.get('exp') || '0', 10);

  if (token && expiry) {
    await chrome.storage.local.set({ proxyToken: { token, expiry } });
    console.log('[SAP Hours Agent] Proxy token stored, closing login tab');
    chrome.tabs.remove(tabId);
  }
});
