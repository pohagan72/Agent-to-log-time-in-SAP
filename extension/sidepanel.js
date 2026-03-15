const chat = document.getElementById('chat');
const userInput = document.getElementById('userInput');
const sendBtn = document.getElementById('sendBtn');
const statusEl = document.getElementById('status');
const minimizeBtn = document.getElementById('minimizeBtn');

let conversationHistory = [];

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

    window.parent.postMessage({
      source: 'sap-hours-agent',
      id,
      type,
      payload,
    }, '*');
  });
}

// Listen for responses from content script
window.addEventListener('message', (event) => {
  if (!event.data || event.data.source !== 'sap-hours-agent-response') return;

  // Handle save-before-reload signal
  if (event.data.type === 'SAVE_BEFORE_RELOAD') {
    saveState();
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
      if (!data.agentState) { resolve(false); return; }
      const state = data.agentState;
      // Only restore if saved within last 5 minutes
      if (Date.now() - state.savedAt > 5 * 60 * 1000) {
        chrome.storage.local.remove('agentState');
        resolve(false);
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
      resolve(true);
    });
  });
}

// --- Init ---

async function init() {
  setStatus('connected', 'Ready');

  // Hide minimize button if not embedded
  if (!isEmbedded && minimizeBtn) {
    minimizeBtn.style.display = 'none';
  }

  const restored = await restoreState();
  if (restored) {
    addMessage('system', 'Page refreshed — entries submitted. Conversation restored.');
  } else {
    const firstName = USER_CONFIG.displayName.split(' ')[0] || 'there';
    addMessage('agent', `Hi ${firstName}! I'm your SAP time entry assistant.\n\nTell me what you worked on last week and I'll fill in your timesheet. For example:\n\n"I spent all week on Agentics AI, 8 hours a day"\n\n"Last week was 4h Agentics AI and 4h AI Platform every day, and I took Friday off"`);
  }
}

// --- Chat UI ---

function formatAgentText(text) {
  // Escape HTML first
  let html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // Bold: **text**
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');

  // Markdown tables: detect lines starting with |
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
        // Collapse multiple blank lines — only add break if previous wasn't empty
        if (out.length > 0 && !out[out.length - 1].endsWith('<br>') && out[out.length - 1] !== '') {
          out.push('<br>');
        }
      } else {
        out.push(`${line}<br>`);
      }
    }
  }
  if (inTable) out.push('</table>');

  // Clean up trailing <br>
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
    } else {
      await processAgentResponse(response);
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

// Fallback for popup mode
async function findSAPTab() {
  const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (activeTab && activeTab.url && activeTab.url.includes(USER_CONFIG.sapHostname)) {
    return activeTab;
  }
  const sapTabs = await chrome.tabs.query({ url: `https://${USER_CONFIG.sapHostname}/*` });
  return sapTabs.length > 0 ? sapTabs[0] : null;
}

// --- Claude API ---

async function callClaude(userMessage, sapState) {
  if (userMessage) {
    conversationHistory.push({ role: 'user', content: userMessage });
  }

  if (!sapState) {
    sapState = await getSAPState();
  }

  const today = new Date().toISOString().split('T')[0];
  const dayOfWeek = new Date().toLocaleDateString('en-US', { weekday: 'long' });

  const favProjects = sapState.favoriteProjects || [];

  let projectSection = '';
  if (favProjects.length > 0) {
    projectSection += 'FAVORITE PROJECTS (read from SAP page):\n';
    favProjects.forEach((p, i) => {
      if (typeof p === 'object') {
        projectSection += `${i + 1}. ${p.project} (Activity: ${p.activity})\n`;
      } else {
        projectSection += `${i + 1}. ${p}\n`;
      }
    });
  }

  const systemPrompt = `You are an AI assistant embedded in a browser extension on the SAP Fiori Time Entry page. You help ${USER_CONFIG.displayName} enter their weekly hours.

CONTEXT:
- Today is ${today} (${dayOfWeek})
- ${USER_CONFIG.displayName} works at ${USER_CONFIG.company}, Personnel #${USER_CONFIG.persNumber}
- Cost Center: ${USER_CONFIG.costCenter}
- Default Role: ${USER_CONFIG.defaultRole}
- Standard work day: 8 hours, Mon-Fri
- Last week = Mon-Fri of the week before today's week

${projectSection || 'No favorites loaded from page.\n'}
PROJECT SEARCH:
You have access to SAP's full project database (thousands of projects). To search for a project, include a search action:

\`\`\`json
{"action": "SEARCH_PROJECTS", "query": "conference"}
\`\`\`

This will search both project IDs and descriptions. Use this when:
- The user mentions an activity that doesn't clearly match a favorite
- You want to find the right project for a specific type of work
- The user asks what projects are available

After finding the right project, you can look up its available activities:

\`\`\`json
{"action": "GET_ACTIVITIES", "projectId": "ADM.000022", "role": "ZADMIN"}
\`\`\`

LOOK UP RECORDED HOURS:
You can look up what hours have already been recorded for any date range:

\`\`\`json
{"action": "GET_RECORDED_HOURS", "startDate": "2026-02-10", "endDate": "2026-02-14"}
\`\`\`

This returns all time entries with project, activity, hours, status (Approved/Pending), and description. Use this when:
- The user asks what they recorded on a specific date or week
- The user wants to check if hours are already entered before adding more
- The user asks about their time entry history or status

CURRENT SAP PAGE STATE:
Day tabs: ${JSON.stringify(sapState.dayTabs || [])}
Current entries: ${JSON.stringify(sapState.currentEntries || [])}
Total hours: ${JSON.stringify(sapState.totalHours)}

WORKFLOW:
1. User describes their week
2. Match activities to favorites if possible
3. If unsure, SEARCH for the right project using the search action
4. Once you have the right project, GET its activities to pick the right activity code
5. Propose the time entries with a clear summary
6. On user confirmation, submit with ENTER_TIME

When ready to submit, include a JSON block with this exact format:

\`\`\`json
{"action": "ENTER_TIME", "entries": [
  {"date": "YYYY-MM-DD", "projectId": "DEV.000982", "projectName": "Agentics AI - 2026", "activity": "Coding", "hours": 8, "description": "Worked on Agentics AI coding"},
  ...
]}
\`\`\`

IMPORTANT: Every entry MUST include a "description" field (SAP rejects entries without comments). Use a brief, professional description of the work.
\`\`\`

RULES:
- Always confirm the plan with the user BEFORE including the ENTER_TIME action
- When proposing entries, ALWAYS show a clear table/list that includes the date, project, hours, AND the description you plan to use for each entry. The description is a required field in SAP and will be visible to managers, so the user must verify it.
- If the user says "yes", "do it", "go ahead", "submit", "looks good", etc., THEN include the ENTER_TIME JSON
- Use SEARCH_PROJECTS proactively when the user mentions non-obvious work (conferences, training, client work, etc.)
- Days should total 8 hours unless the user says otherwise
- If the user says they took a day off, skip that day entirely (PTO is handled separately in SAP)
- Generate professional, accurate descriptions based on what the user told you (e.g. "Agentics AI development and coding", "AI Platform requirements and design")
- Be concise and friendly
- If unsure about anything, ask`;

  try {
    const resp = await fetch(CONFIG.endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': CONFIG.apiKey,
        'anthropic-version': CONFIG.apiVersion,
      },
      body: JSON.stringify({
        model: CONFIG.model,
        max_tokens: 1024,
        system: systemPrompt,
        messages: conversationHistory,
      }),
    });

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
            `${e.date}: ${e.hours}h — ${e.project} (${e.projectName}) [${e.status}]`
          ).join('\n');
          addMessage('system', `Found ${entries.length} entries (${result.totalHours}h total):\n${summary}`);
        }

        const hoursMsg = entries.length > 0
          ? `Recorded hours from ${action.startDate} to ${action.endDate} (${result.totalHours}h total):\n${entries.map(e => `- ${e.date}: ${e.hours}h on ${e.project} / ${e.projectName}, activity: ${e.activityName || e.activity}, status: ${e.status}${e.description ? ', desc: ' + e.description : ''}`).join('\n')}`
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
      }
    } catch (e) {
      addMessage('agent', text);
    }
  } else {
    addMessage('agent', text);
  }
}

async function executeTimeEntries(entries) {
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

  addMessage('system', 'Done! Please review the entries in SAP.');
}

// --- Start ---
init();
