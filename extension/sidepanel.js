const chat = document.getElementById('chat');
const userInput = document.getElementById('userInput');
const sendBtn = document.getElementById('sendBtn');
const statusEl = document.getElementById('status');
const minimizeBtn = document.getElementById('minimizeBtn');

let conversationHistory = [];
let userConfig = null;

// Detect if running inside an iframe (embedded in SAP page) vs standalone popup
const isEmbedded = window.parent !== window;

// --- PostMessage Communication (embedded mode) ---

let pendingMessages = {};

function postMessageToParent(type, payload) {
  return new Promise((resolve, reject) => {
    const id = `${Date.now()}-${Math.random()}`;
    const timeout = setTimeout(() => {
      delete pendingMessages[id];
      reject(new Error(`Timeout waiting for ${type} response`));
    }, 30000);

    pendingMessages[id] = { resolve, reject, timeout };

    const targetOrigin = `https://${CONFIG.sapHostname}`;
    window.parent.postMessage({
      source: 'sap-hours-agent',
      id,
      type,
      payload,
    }, targetOrigin);
  });
}

// Listen for responses from content script (SAP page origin only)
const EXPECTED_ORIGIN = `https://${CONFIG.sapHostname}`;
window.addEventListener('message', (event) => {
  if (event.origin !== EXPECTED_ORIGIN) return;
  if (!event.data || event.data.source !== 'sap-hours-agent-response') return;

  // Handle save-before-reload signal
  if (event.data.type === 'SAVE_BEFORE_RELOAD') {
    saveState();
    return;
  }

  // Handle user config push from content script auto-discovery
  if (event.data.type === 'USER_CONFIG_READY') {
    if (!userConfig && event.data.config) {
      userConfig = event.data.config;
      console.log('[SAP Hours Agent] Received user config from content script:', userConfig.displayName);
    }
    return;
  }

  const { id, result } = event.data;
  const pending = pendingMessages[id];
  if (pending) {
    clearTimeout(pending.timeout);
    delete pendingMessages[id];
    pending.resolve(result);
  }
});

// --- Load User Config ---

async function loadUserConfig() {
  // 1. Check chrome.storage cache
  const cached = await new Promise(resolve => {
    chrome.storage.local.get('sapUserConfig', data => resolve(data.sapUserConfig));
  });

  if (cached && cached.persNumber) {
    console.log('[SAP Hours Agent] Loaded user config from cache:', cached.displayName);
    return cached;
  }

  // 2. Request from content script (embedded mode)
  if (isEmbedded) {
    try {
      const config = await postMessageToParent('GET_USER_CONFIG', {});
      if (config && config.persNumber) {
        console.log('[SAP Hours Agent] Got user config from content script:', config.displayName);
        return config;
      }
    } catch (e) {
      console.log('[SAP Hours Agent] Could not get config from content script:', e.message);
    }
  } else {
    // 3. Request via chrome.tabs message (popup mode)
    try {
      const tab = await findSAPTab();
      if (tab) {
        const config = await chrome.tabs.sendMessage(tab.id, { type: 'GET_USER_CONFIG' });
        if (config && config.persNumber) {
          console.log('[SAP Hours Agent] Got user config from SAP tab:', config.displayName);
          return config;
        }
      }
    } catch (e) {
      console.log('[SAP Hours Agent] Could not get config from SAP tab:', e.message);
    }
  }

  // 4. Return defaults if discovery hasn't completed yet
  console.warn('[SAP Hours Agent] User config not yet available, using defaults');
  return {
    persNumber: '',
    userName: '',
    displayName: '',
    userTitle: '',
    company: '',
    costCenter: '',
    defaultRole: 'ZADMIN',
    sapHostname: CONFIG.sapHostname,
  };
}

// --- State Persistence ---

function saveState() {
  const chatMessages = [];
  chat.querySelectorAll('.message').forEach(div => {
    const type = ['agent', 'user', 'system', 'error'].find(t => div.classList.contains(t)) || 'agent';
    chatMessages.push({ type, text: type === 'agent' ? div.innerHTML : div.textContent });
  });
  const state = { chatMessages, conversationHistory, savedAt: Date.now() };
  chrome.storage.local.set({ agentState: state });
  console.log('[SAP Hours Agent] State saved:', chatMessages.length, 'messages');
}

async function restoreState() {
  return new Promise(resolve => {
    chrome.storage.local.get('agentState', (data) => {
      if (!data.agentState) { resolve(null); return; }
      const state = data.agentState;
      // Only restore if saved within last 5 minutes
      if (Date.now() - state.savedAt > 5 * 60 * 1000) {
        chrome.storage.local.remove('agentState');
        resolve(null);
        return;
      }
      // Restore chat UI
      for (const msg of state.chatMessages) {
        const div = document.createElement('div');
        div.className = `message ${msg.type}`;
        if (msg.type === 'agent') {
          div.innerHTML = msg.text;
        } else {
          div.textContent = msg.text;
        }
        chat.appendChild(div);
      }
      chat.scrollTop = chat.scrollHeight;
      // Restore conversation history
      conversationHistory = state.conversationHistory || [];
      // Clear saved state
      chrome.storage.local.remove('agentState');
      console.log('[SAP Hours Agent] State restored:', state.chatMessages.length, 'messages');
      resolve(state.pendingEntries || true);
    });
  });
}

// --- Init ---

async function init() {
  setStatus('connected', 'Discovering...');

  // Hide minimize button if not embedded
  if (!isEmbedded && minimizeBtn) {
    minimizeBtn.style.display = 'none';
  }

  // Load user config (auto-discovered from SAP)
  userConfig = await loadUserConfig();

  if (!userConfig.persNumber) {
    setStatus('error', 'Setup needed');
    addMessage('system', 'Could not auto-detect your SAP user info. Make sure you are on the SAP Time Entry page and refresh.');
  } else {
    setStatus('connected', 'Ready');
  }

  const restored = await restoreState();
  if (restored && Array.isArray(restored)) {
    // Navigated here to submit pending entries — proceed automatically
    addMessage('system', 'Conversation restored. Submitting entries...');
    await executeTimeEntries(restored);
  } else if (restored) {
    addMessage('system', 'Page refreshed — entries submitted. Conversation restored.');
  } else {
    await showWelcome();
  }
}

async function showWelcome() {
  const firstName = (userConfig.displayName || '').split(' ')[0] || 'there';

  // Try to get context for a smarter greeting
  let weekInfo = '';
  let favExamples = '';
  let favs = [];

  try {
    const weekTotal = await getWeekTotal();
    if (weekTotal && !weekTotal.error) {
      const today = new Date();
      const dayOfWeek = today.getDay(); // 0=Sun, 1=Mon, ...
      const mon = new Date(today);
      mon.setDate(today.getDate() - ((dayOfWeek + 6) % 7));
      const sun = new Date(mon);
      sun.setDate(mon.getDate() + 6);
      const fmt = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      weekInfo = `\nYou have **${weekTotal.hours}h** logged this week (${fmt(mon)}–${fmt(sun)}).`;
    }
  } catch (e) { /* ignore */ }

  try {
    favs = await getFavorites();
    if (favs.length >= 2) {
      const f1 = favs[0].projectDesc || favs[0].project;
      const f2 = favs[1].projectDesc || favs[1].project;
      favExamples = `\n\n"I spent all week on ${f1}, 8 hours a day"\n\n"Last week was 4h ${f1} and 4h ${f2} every day, and I took Friday off"`;
    } else if (favs.length === 1) {
      const f1 = favs[0].projectDesc || favs[0].project;
      favExamples = `\n\n"I spent all week on ${f1}, 8 hours a day"\n\n"Last week was 8h ${f1} every day, I took Friday off"`;
    }
  } catch (e) { /* ignore */ }

  if (!favExamples) {
    favExamples = '\n\n"I spent all week on project X, 8 hours a day"\n\n"Last week was 4h project A and 4h project B every day, and I took Friday off"';
  }

  addMessage('agent', `Hi ${firstName}! I'm your SAP time entry assistant.${weekInfo}\n\nTell me what you worked on and I'll fill in your timesheet. For example:${favExamples}`);

  // Update quick action buttons with user's favorites
  const quickActions = document.getElementById('quickActions');
  if (quickActions && favs && favs.length > 0) {
    const f1 = favs[0].projectDesc || favs[0].project;
    const buttons = [
      { label: `8h ${shortName(f1)} (last week)`, msg: `Enter 8 hours on ${f1} for all of last week` },
      { label: 'Missing days', msg: 'Which days am I missing hours this month?' },
    ];
    if (favs.length >= 2) {
      const f2 = favs[1].projectDesc || favs[1].project;
      buttons.push({ label: `Split ${shortName(f1)}/${shortName(f2)}`, msg: `Split last week: 4h ${f1} and 4h ${f2} each day` });
    } else {
      buttons.push({ label: 'Copy last week', msg: 'Copy last week to this week' });
    }
    quickActions.innerHTML = '';
    for (const b of buttons) {
      const btn = document.createElement('button');
      btn.textContent = b.label;
      btn.dataset.msg = b.msg;
      btn.addEventListener('click', () => { userInput.value = b.msg; handleSend(); });
      quickActions.appendChild(btn);
    }
  }
}

function shortName(name) {
  // Trim project names for button labels: "Agentics AI - 2026" -> "Agentics AI"
  return name.replace(/\s*-\s*\d{4}$/, '').replace(/^LS-Dev \d+-/, '').trim();
}

// --- Chat UI ---

function formatAgentText(text) {
  let html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // Bold: **text**
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');

  // Markdown tables
  const lines = html.split('\n');
  let inTable = false;
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.startsWith('|') && line.endsWith('|')) {
      if (/^\|[\s\-:|]+\|$/.test(line)) continue;
      const cells = line.split('|').filter(c => c.trim() !== '').map(c => c.trim());
      if (!inTable) {
        out.push('<table>');
        out.push('<tr>' + cells.map(c => `<th>${c}</th>`).join('') + '</tr>');
        inTable = true;
      } else {
        out.push('<tr>' + cells.map(c => `<td>${c}</td>`).join('') + '</tr>');
      }
    } else {
      if (inTable) { out.push('</table>'); inTable = false; }
      if (line === '') {
        if (out.length > 0 && !out[out.length - 1].endsWith('<br>') && out[out.length - 1] !== '') {
          out.push('<br>');
        }
      } else {
        out.push(`${line}<br>`);
      }
    }
  }
  if (inTable) out.push('</table>');

  let result = out.join('');
  result = result.replace(/(<br>)+$/, '');
  return result;
}

function addMessage(type, text) {
  const div = document.createElement('div');
  div.className = `message ${type}`;
  if (type === 'agent') {
    div.innerHTML = formatAgentText(text);
  } else {
    div.textContent = text;
  }
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
  return div;
}

function showTyping() {
  const div = document.createElement('div');
  div.className = 'typing';
  div.id = 'typing';
  div.textContent = 'Thinking';
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
}

function hideTyping() {
  const el = document.getElementById('typing');
  if (el) el.remove();
}

function setStatus(cls, text) {
  statusEl.className = `status ${cls}`;
  statusEl.textContent = text;
}

// --- Minimize Button ---

if (minimizeBtn) {
  minimizeBtn.addEventListener('click', () => {
    if (isEmbedded) {
      postMessageToParent('MINIMIZE_PANEL', {}).catch(() => {});
    }
  });
}

// --- Quick Actions ---

document.querySelectorAll('.quick-actions button').forEach((btn) => {
  btn.addEventListener('click', () => {
    userInput.value = btn.dataset.msg;
    handleSend();
  });
});

// --- Send Message ---

sendBtn.addEventListener('click', handleSend);
userInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    handleSend();
  }
});

async function handleSend() {
  const text = userInput.value.trim();
  if (!text) return;

  addMessage('user', text);
  userInput.value = '';
  sendBtn.disabled = true;
  showTyping();
  setStatus('connected', 'Thinking...');

  try {
    const sapState = await getSAPState();
    const response = await callClaude(text, sapState);
    hideTyping();

    if (response.error) {
      addMessage('error', response.error);
      setStatus('error', 'Error');
    } else if (response.text) {
      await processAgentResponse(response);
      setStatus('connected', 'Ready');
    } else {
      // null error = auth redirect in progress, nothing to display
      setStatus('connected', 'Ready');
    }
  } catch (err) {
    hideTyping();
    addMessage('error', `Error: ${err.message}`);
    setStatus('error', 'Error');
  }

  sendBtn.disabled = false;
  userInput.focus();
}

// --- SAP Communication Layer ---

async function getSAPState() {
  try {
    if (isEmbedded) {
      const response = await postMessageToParent('GET_STATE', {});
      return { onSAPPage: true, ...response };
    } else {
      const tab = await findSAPTab();
      if (!tab) return { onSAPPage: false };
      const response = await chrome.tabs.sendMessage(tab.id, { type: 'GET_STATE' });
      return { onSAPPage: true, ...response };
    }
  } catch (e) {
    console.error('[SAP Hours Agent] getSAPState failed:', e);
    return { onSAPPage: false, error: e.message };
  }
}

async function sendSAPAction(action) {
  if (isEmbedded) {
    return postMessageToParent('EXECUTE_ACTION', { action });
  }
  const tab = await findSAPTab();
  if (!tab) throw new Error('No SAP tab found.');
  return chrome.tabs.sendMessage(tab.id, { type: 'EXECUTE_ACTION', action });
}

async function navigateToTimeEntry() {
  if (isEmbedded) {
    return postMessageToParent('NAVIGATE_TIME_ENTRY', {});
  }
  const tab = await findSAPTab();
  if (!tab) throw new Error('No SAP tab found.');
  const timeEntryUrl = `https://${CONFIG.sapHostname}/sap/bc/ui2/flp#FioriTime-Enter`;
  await chrome.tabs.update(tab.id, { url: timeEntryUrl });
}

async function searchSAPProjects(query) {
  if (isEmbedded) {
    return postMessageToParent('SEARCH_PROJECTS', { query });
  }
  const tab = await findSAPTab();
  if (!tab) return { projects: [], error: 'No SAP tab found' };
  return chrome.tabs.sendMessage(tab.id, { type: 'SEARCH_PROJECTS', query });
}

async function getRecordedHours(startDate, endDate) {
  if (isEmbedded) {
    return postMessageToParent('GET_RECORDED_HOURS', { startDate, endDate });
  }
  const tab = await findSAPTab();
  if (!tab) return { entries: [], totalHours: 0, error: 'No SAP tab found' };
  return chrome.tabs.sendMessage(tab.id, { type: 'GET_RECORDED_HOURS', startDate, endDate });
}

async function getSAPProjectActivities(projectId, role) {
  if (isEmbedded) {
    return postMessageToParent('GET_PROJECT_ACTIVITIES', { projectId, role });
  }
  const tab = await findSAPTab();
  if (!tab) return { activities: [], error: 'No SAP tab found' };
  return chrome.tabs.sendMessage(tab.id, { type: 'GET_PROJECT_ACTIVITIES', projectId, role });
}

async function getFavorites() {
  if (isEmbedded) {
    return postMessageToParent('GET_FAVORITES', {});
  }
  const tab = await findSAPTab();
  if (!tab) return [];
  return chrome.tabs.sendMessage(tab.id, { type: 'GET_FAVORITES' });
}

async function getCalendarStatus(referenceDate) {
  if (isEmbedded) {
    return postMessageToParent('GET_CALENDAR_STATUS', { referenceDate });
  }
  const tab = await findSAPTab();
  if (!tab) return { days: [], error: 'No SAP tab found' };
  return chrome.tabs.sendMessage(tab.id, { type: 'GET_CALENDAR_STATUS', referenceDate });
}

async function getWeekTotal() {
  if (isEmbedded) {
    return postMessageToParent('GET_WEEK_TOTAL', {});
  }
  const tab = await findSAPTab();
  if (!tab) return { hours: 0, error: 'No SAP tab found' };
  return chrome.tabs.sendMessage(tab.id, { type: 'GET_WEEK_TOTAL' });
}

// Fallback for popup mode
async function findSAPTab() {
  const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (activeTab && activeTab.url && activeTab.url.includes(CONFIG.sapHostname)) {
    return activeTab;
  }
  const sapTabs = await chrome.tabs.query({ url: `https://${CONFIG.sapHostname}/*` });
  return sapTabs.length > 0 ? sapTabs[0] : null;
}

// --- Easy Auth login (opens proxy login page in a new tab) ---

// --- Proxy token management ---
// After Easy Auth login, /api/token mints a signed token and posts it back via
// window.opener.postMessage. We cache it in chrome.storage (8h TTL).

let cachedProxyToken = null;
let cachedProxyTokenExpiry = 0;

async function getProxyToken() {
  // Return in-memory cache first
  if (cachedProxyToken && Date.now() < cachedProxyTokenExpiry - 60000) {
    return cachedProxyToken;
  }
  // Try chrome.storage
  const stored = await new Promise(resolve =>
    chrome.storage.local.get('proxyToken', d => resolve(d.proxyToken))
  );
  if (stored && stored.token && Date.now() < stored.expiry - 60000) {
    cachedProxyToken = stored.token;
    cachedProxyTokenExpiry = stored.expiry;
    return cachedProxyToken;
  }
  return null;
}

async function saveProxyToken(token, expiry) {
  cachedProxyToken = token;
  cachedProxyTokenExpiry = expiry;
  await new Promise(resolve => chrome.storage.local.set({ proxyToken: { token, expiry } }, resolve));
}


async function triggerEasyAuthLogin() {
  await chrome.tabs.create({ url: 'https://sap-hours-proxy.azurewebsites.net/api/token' });
}

// --- Claude API ---

async function callClaude(userMessage, sapState) {
  if (userMessage) {
    conversationHistory.push({ role: 'user', content: userMessage });
  }

  if (!sapState) {
    sapState = await getSAPState();
  }

  // Refresh user config if we don't have it yet
  if (!userConfig || !userConfig.persNumber) {
    userConfig = await loadUserConfig();
  }

  const favProjects = sapState.favoriteProjects || [];

  // Build context for server-side system prompt (no prompt sent from client)
  const context = {
    displayName: userConfig.displayName || 'User',
    company: userConfig.company || 'Unknown',
    persNumber: userConfig.persNumber || 'Unknown',
    costCenter: userConfig.costCenter || '',
    costCenterName: userConfig.costCenterName || '',
    defaultRole: userConfig.defaultRole || 'ZADMIN',
    favorites: favProjects,
    sapState: {
      dayTabs: sapState.dayTabs || [],
      currentEntries: sapState.currentEntries || [],
      totalHours: sapState.totalHours,
    },
  };

  try {
    let token = await getProxyToken();
    if (!token) {
      addMessage('system', 'Signing in...');
      await triggerEasyAuthLogin();
      for (let i = 0; i < 30; i++) {
        await new Promise(resolve => setTimeout(resolve, 1000));
        token = await getProxyToken();
        if (token) break;
      }
      if (!token) {
        addMessage('error', 'Sign-in timed out. Please try again.');
        return { error: null };
      }
    }

    const resp = await fetch(CONFIG.proxyEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Agent-Token': token,
      },
      body: JSON.stringify({
        messages: conversationHistory,
        context: context,
      }),
    });

    if (resp.status === 401) {
      cachedProxyToken = null;
      cachedProxyTokenExpiry = 0;
      await new Promise(resolve => chrome.storage.local.remove('proxyToken', resolve));
      addMessage('system', 'Session expired. Opening login page — after signing in, come back and try again.');
      await triggerEasyAuthLogin();
      return { error: null };
    }

    if (!resp.ok) {
      const errText = await resp.text();
      return { error: `API error (${resp.status}): ${errText}` };
    }

    const data = await resp.json();
    const assistantMessage = data.content[0].text;
    conversationHistory.push({ role: 'assistant', content: assistantMessage });
    return { text: assistantMessage };
  } catch (err) {
    return { error: `Network error: ${err.message}` };
  }
}

// --- Process Agent Response ---

async function processAgentResponse(response) {
  const text = response.text;

  const jsonBlockMatch = text.match(/```json\s*([\s\S]*?)```/);
  let action = null;

  if (jsonBlockMatch) {
    try {
      action = JSON.parse(jsonBlockMatch[1].trim());
    } catch (e) { /* JSON parse failed */ }
  }

  if (!action) {
    const rawJsonMatch = text.match(/\{"action"\s*:\s*"[A-Z_]+[\s\S]*?\}(?:\s*\]?\s*\})?/);
    if (rawJsonMatch) {
      try {
        action = JSON.parse(rawJsonMatch[0]);
      } catch (e) { /* ignore */ }
    }
  }

  if (action && action.action) {
    try {
      const displayText = text
        .replace(/```json[\s\S]*?```/g, '')
        .replace(/\{"action"\s*:\s*"[A-Z_]+[\s\S]*?\}(?:\s*\]?\s*\})?/g, '')
        .trim();
      if (displayText) addMessage('agent', displayText);

      if (action.action === 'ENTER_TIME') {
        await executeTimeEntries(action.entries);

      } else if (action.action === 'SEARCH_PROJECTS') {
        addMessage('system', `Searching SAP for "${action.query}"...`);
        const result = await searchSAPProjects(action.query);
        const projects = result.projects || [];

        if (result.error) {
          addMessage('error', `SAP search error: ${result.error}`);
        } else if (projects.length === 0) {
          addMessage('system', `No projects found for "${action.query}".`);
        } else {
          const summary = projects.slice(0, 20).map(p => `${p.id} — ${p.name}`).join('\n');
          addMessage('system', `Found ${projects.length} projects:\n${summary}`);
        }

        const searchResultMsg = projects.length > 0
          ? `SAP project search results for "${action.query}":\n${projects.slice(0, 20).map(p => `- ${p.id}: ${p.name}`).join('\n')}`
          : `No SAP projects found matching "${action.query}". Try a different search term.`;

        conversationHistory.push({ role: 'user', content: `[SYSTEM: ${searchResultMsg}]` });

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'GET_ACTIVITIES') {
        addMessage('system', `Looking up activities for ${action.projectId}...`);
        const result = await getSAPProjectActivities(action.projectId, action.role || 'ZADMIN');
        const activities = result.activities || [];

        if (activities.length === 0) {
          addMessage('system', `No activities found for ${action.projectId}.`);
        } else {
          const summary = activities.map(a => `${a.code}: ${a.description}`).join('\n');
          addMessage('system', `Activities for ${action.projectId}:\n${summary}`);
        }

        const actResultMsg = activities.length > 0
          ? `Activities for project ${action.projectId}:\n${activities.map(a => `- ${a.code}: ${a.description}`).join('\n')}`
          : `No activities found for ${action.projectId}.`;

        conversationHistory.push({ role: 'user', content: `[SYSTEM: ${actResultMsg}]` });

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'GET_RECORDED_HOURS') {
        addMessage('system', `Looking up hours from ${action.startDate} to ${action.endDate}...`);
        const result = await getRecordedHours(action.startDate, action.endDate);
        const entries = result.entries || [];

        if (result.error) {
          addMessage('error', `SAP API error: ${result.error}`);
        } else if (entries.length === 0) {
          addMessage('system', `No hours recorded for ${action.startDate} to ${action.endDate}.`);
        } else {
          const summary = entries.map(e =>
            `${e.date} (${e.dayOfWeek}): ${e.hours}h — ${e.project} (${e.projectName}) [${e.status}]`
          ).join('\n');
          addMessage('system', `Found ${entries.length} entries (${result.totalHours}h total):\n${summary}`);
        }

        const hoursMsg = entries.length > 0
          ? `Recorded hours from ${action.startDate} to ${action.endDate} (${result.totalHours}h total):\n${entries.map(e => `- ${e.date} (${e.dayOfWeek}): ${e.hours}h on ${e.project} / ${e.projectName}, activity: ${e.activityName || e.activity}, counter: ${e.counter}, status: ${e.status}${e.description ? ', desc: ' + e.description : ''}`).join('\n')}`
          : `No hours recorded from ${action.startDate} to ${action.endDate}.`;

        conversationHistory.push({ role: 'user', content: `[SYSTEM: ${hoursMsg}]` });

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'GET_CALENDAR_STATUS') {
        addMessage('system', `Checking calendar for ${action.referenceDate}...`);
        const result = await getCalendarStatus(action.referenceDate);
        const days = result.days || [];

        if (result.error) {
          addMessage('error', `Calendar error: ${result.error}`);
        }

        const calMsg = days.length > 0
          ? `Calendar status for month of ${action.referenceDate}:\n${days.map(d => `- ${d.date} (${d.dayOfWeek}): ${d.hours}h ${d.complete ? '(complete)' : '(INCOMPLETE)'}`).join('\n')}\nDays not listed have NO hours entered.`
          : `No hours recorded in the month of ${action.referenceDate}.`;

        addMessage('system', calMsg);
        conversationHistory.push({ role: 'user', content: `[SYSTEM: ${calMsg}]` });

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'GET_WEEK_TOTAL') {
        const result = await getWeekTotal();
        const msg = result.error
          ? `Could not get week total: ${result.error}`
          : `Current week: ${result.hours}h (${result.label})`;
        addMessage('system', msg);
        conversationHistory.push({ role: 'user', content: `[SYSTEM: ${msg}]` });

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'DELETE_ENTRY') {
        addMessage('system', `Deleting entry ${action.counter}...`);
        const result = await sendSAPAction({ type: 'DELETE_ENTRY', counter: action.counter });
        if (result.success) {
          addMessage('system', result.message);
          conversationHistory.push({ role: 'user', content: `[SYSTEM: Entry ${action.counter} deleted successfully.]` });
        } else {
          addMessage('error', `Delete failed: ${result.error}`);
          conversationHistory.push({ role: 'user', content: `[SYSTEM: Delete failed: ${result.error}]` });
        }

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'COPY_WEEK') {
        const verb = action.move ? 'Moving' : 'Copying';
        addMessage('system', `${verb} entries from ${action.fromDate} to ${action.toDate}...`);
        const result = await sendSAPAction({
          type: 'COPY_WEEK',
          fromDate: action.fromDate,
          toDate: action.toDate,
          move: action.move || false,
        });
        if (result.success) {
          addMessage('system', result.message || `${verb} complete!`);
          conversationHistory.push({ role: 'user', content: `[SYSTEM: ${verb} from ${action.fromDate} to ${action.toDate} succeeded.]` });
        } else {
          addMessage('error', `${verb} failed: ${result.error}`);
          conversationHistory.push({ role: 'user', content: `[SYSTEM: ${verb} failed: ${result.error}]` });
        }

        setStatus('connected', 'Thinking...');
        showTyping();
        const followUp = await callClaude('', null);
        hideTyping();
        if (followUp.error) {
          addMessage('error', followUp.error);
        } else {
          await processAgentResponse(followUp);
        }

      } else if (action.action === 'ADD_FAVORITE' || action.action === 'REMOVE_FAVORITE') {
        const isAdd = action.action === 'ADD_FAVORITE';
        const verb = isAdd ? 'Adding' : 'Removing';
        addMessage('system', `${verb} favorite ${action.projectId}/${action.activity}...`);
        const result = await sendSAPAction({
          type: isAdd ? 'ADD_FAVORITE' : 'REMOVE_FAVORITE',
          projectId: action.projectId,
          activity: action.activity,
          description: action.description || '',
        });
        if (result.success) {
          addMessage('system', result.message || `Favorite ${isAdd ? 'added' : 'removed'}!`);
        } else {
          addMessage('error', `Failed: ${result.error}`);
        }
      }
    } catch (e) {
      addMessage('agent', text);
    }
  } else {
    addMessage('agent', text);
  }
}

async function executeTimeEntries(entries) {
  const currentState = await getSAPState();
  if (currentState.page !== 'timeEntry') {
    addMessage('system', 'Navigating to Time Entry...');
    // Save chat + conversation + pending entries before navigation destroys the iframe
    const chatMessages = [];
    chat.querySelectorAll('.message').forEach(div => {
      const type = ['agent', 'user', 'system', 'error'].find(t => div.classList.contains(t)) || 'agent';
      chatMessages.push({ type, text: type === 'agent' ? div.innerHTML : div.textContent });
    });
    const state = { chatMessages, conversationHistory, savedAt: Date.now(), pendingEntries: entries };
    await new Promise(resolve => chrome.storage.local.set({ agentState: state }, resolve));
    try {
      await navigateToTimeEntry();
    } catch (err) {
      addMessage('error', `Could not navigate to Time Entry: ${err.message}`);
    }
    return;
  }

  addMessage('system', `Entering ${entries.length} time entries into SAP...`);

  const byDate = {};
  for (const entry of entries) {
    if (!byDate[entry.date]) byDate[entry.date] = [];
    byDate[entry.date].push(entry);
  }

  for (const [date, dayEntries] of Object.entries(byDate)) {
    const dayName = new Date(date + 'T12:00:00').toLocaleDateString('en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
    });

    addMessage('system', `Processing ${dayName}...`);

    try {
      const result = await sendSAPAction({
        type: 'ENTER_DAY',
        date,
        entries: dayEntries,
      });

      console.log(`[SAP Hours Agent] ENTER_DAY result for ${date}:`, result);

      if (!result) {
        addMessage('error', `${dayName}: No response from SAP page.`);
      } else if (result.success) {
        addMessage('system', `${dayName}: ${result.message || 'Done'}`);
      } else {
        addMessage('error', `${dayName}: ${result.error || result.message || 'Unknown error'}`);
      }
    } catch (err) {
      addMessage('error', `${dayName}: Failed - ${err.message}`);
    }
  }

  addMessage('system', 'Done! SAP page will refresh to show new entries.');
  saveState();
  if (isEmbedded) {
    postMessageToParent('RELOAD_PAGE', {}).catch(() => {});
  }
}

// --- Start ---
init();
