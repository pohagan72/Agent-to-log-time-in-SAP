// SAP Fiori Time Entry - Content Script
// USER_CONFIG is loaded before this script via manifest.json content_scripts order

console.log('[SAP Hours Agent] Content script loaded');

const SAP_ODATA_BASE = USER_CONFIG.sapODataPath;
const SAP_CLIENT = `sap-client=${USER_CONFIG.sapClient}`;
const PERS_NUMBER = USER_CONFIG.persNumber;
const USER_NAME = USER_CONFIG.userName;

// --- Inject Agent Panel into SAP page ---

let panelVisible = true;

function injectAgentPanel() {
  if (document.getElementById('sap-hours-agent-root')) return;

  // Container
  const root = document.createElement('div');
  root.id = 'sap-hours-agent-root';
  root.style.cssText = `
    position: fixed;
    top: 0;
    right: 0;
    height: 100vh;
    z-index: 999999;
    display: flex;
    flex-direction: row;
    pointer-events: none;
  `;

  // Toggle tab (always visible on the edge)
  const toggle = document.createElement('div');
  toggle.id = 'sap-agent-toggle';
  toggle.innerHTML = `<span style="writing-mode: vertical-rl; text-orientation: mixed; font-size: 12px; font-weight: 600; letter-spacing: 1px;">SAP AI</span>`;
  toggle.style.cssText = `
    pointer-events: auto;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    width: 28px;
    background: #e94560;
    color: white;
    border-radius: 8px 0 0 8px;
    padding: 12px 4px;
    align-self: center;
    box-shadow: -2px 0 10px rgba(0,0,0,0.3);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    transition: background 0.2s;
  `;
  toggle.addEventListener('mouseenter', () => { toggle.style.background = '#c73652'; });
  toggle.addEventListener('mouseleave', () => { toggle.style.background = '#e94560'; });

  // Panel container
  const panel = document.createElement('div');
  panel.id = 'sap-agent-panel';
  panel.style.cssText = `
    pointer-events: auto;
    width: 400px;
    height: 100vh;
    border-left: 2px solid #e94560;
    box-shadow: -4px 0 20px rgba(0,0,0,0.4);
    background: #1a1a2e;
    transition: width 0.3s ease, opacity 0.3s ease;
    overflow: hidden;
  `;

  // Iframe loads sidepanel.html from extension
  const iframe = document.createElement('iframe');
  iframe.id = 'sap-agent-iframe';
  iframe.src = chrome.runtime.getURL('sidepanel.html');
  iframe.style.cssText = `
    width: 100%;
    height: 100%;
    border: none;
    background: #1a1a2e;
  `;

  panel.appendChild(iframe);
  root.appendChild(toggle);
  root.appendChild(panel);
  document.body.appendChild(root);

  // Toggle click
  toggle.addEventListener('click', () => {
    togglePanel();
  });

  console.log('[SAP Hours Agent] Panel injected');
}

function togglePanel() {
  const panel = document.getElementById('sap-agent-panel');
  const toggle = document.getElementById('sap-agent-toggle');
  if (!panel) return;

  panelVisible = !panelVisible;
  if (panelVisible) {
    panel.style.width = '400px';
    panel.style.opacity = '1';
    panel.style.pointerEvents = 'auto';
  } else {
    panel.style.width = '0px';
    panel.style.opacity = '0';
    panel.style.pointerEvents = 'none';
  }
}

function minimizePanel() {
  if (panelVisible) togglePanel();
}

// --- Message Handler (from extension background/popup) ---

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  console.log('[SAP Hours Agent] Chrome message:', msg.type);
  if (msg.type === 'TOGGLE_PANEL') {
    togglePanel();
    sendResponse({ ok: true });
    return;
  }
  // Keep existing handlers for backward compatibility (popup mode)
  if (msg.type === 'GET_STATE') {
    getPageState().then(sendResponse).catch(e => sendResponse({ error: e.message }));
    return true;
  }
  if (msg.type === 'EXECUTE_ACTION') {
    executeAction(msg.action).then(sendResponse).catch(e => sendResponse({ success: false, error: e.message }));
    return true;
  }
  if (msg.type === 'SEARCH_PROJECTS') {
    searchProjects(msg.query).then(sendResponse).catch(e => sendResponse({ projects: [], error: e.message }));
    return true;
  }
  if (msg.type === 'GET_PROJECT_ACTIVITIES') {
    getProjectActivities(msg.projectId, msg.role).then(sendResponse).catch(e => sendResponse({ activities: [], error: e.message }));
    return true;
  }
  if (msg.type === 'GET_RECORDED_HOURS') {
    getRecordedHours(msg.startDate, msg.endDate).then(sendResponse).catch(e => sendResponse({ entries: [], totalHours: 0, error: e.message }));
    return true;
  }
});

// --- PostMessage Bridge (from embedded iframe) ---

window.addEventListener('message', async (event) => {
  if (!event.data || event.data.source !== 'sap-hours-agent') return;

  const { id, type, payload } = event.data;
  let result;

  try {
    switch (type) {
      case 'GET_STATE':
        result = await getPageState();
        break;
      case 'SEARCH_PROJECTS':
        result = await searchProjects(payload.query);
        break;
      case 'GET_PROJECT_ACTIVITIES':
        result = await getProjectActivities(payload.projectId, payload.role);
        break;
      case 'GET_RECORDED_HOURS':
        result = await getRecordedHours(payload.startDate, payload.endDate);
        break;
      case 'EXECUTE_ACTION':
        result = await executeAction(payload.action);
        break;
      case 'MINIMIZE_PANEL':
        minimizePanel();
        result = { ok: true };
        break;
      default:
        result = { error: `Unknown message type: ${type}` };
    }
  } catch (e) {
    console.error(`[SAP Hours Agent] postMessage handler error (${type}):`, e);
    result = { error: e.message };
  }

  // Send response back to iframe
  const iframe = document.getElementById('sap-agent-iframe');
  if (iframe && iframe.contentWindow) {
    iframe.contentWindow.postMessage({
      source: 'sap-hours-agent-response',
      id,
      result,
    }, '*');
  }
});

// --- Read SAP Page State ---

async function getPageState() {
  try {
    const url = window.location.href;
    const isTimeEntry = url.includes('FioriTime-Enter');

    if (!isTimeEntry) {
      return {
        page: 'other',
        url,
        message: 'Not on Time Entry page. Navigate to Time Entry first.',
      };
    }

    const dayTabs = readDayTabs();
    const currentEntries = readCurrentEntries();
    const totalHours = readTotalHours();
    const favoriteProjects = readFavoriteProjects();

    return {
      page: 'timeEntry',
      url,
      dayTabs,
      currentEntries,
      totalHours,
      favoriteProjects,
    };
  } catch (err) {
    return { page: 'error', error: err.message };
  }
}

function readDayTabs() {
  const tabs = [];
  const tabItems = document.querySelectorAll('.sapMITBItem');
  tabItems.forEach((item) => {
    const textEl = item.querySelector('.sapMITBText');
    if (textEl) {
      tabs.push({
        text: textEl.textContent.trim(),
        hasAlert: !!item.querySelector('.sapMITBFilterNegative, [src*="alert"]'),
        isSelected: item.classList.contains('sapMITBSelected'),
      });
    }
  });
  return tabs;
}

function readCurrentEntries() {
  const entries = [];
  const rows = document.querySelectorAll('.sapMListItems .sapMLIB, table tbody tr');
  rows.forEach((row) => {
    const cells = row.querySelectorAll('td, .sapMListTblCell');
    if (cells.length >= 2) {
      entries.push({ text: row.textContent.trim().replace(/\s+/g, ' ') });
    }
  });
  return entries;
}

function readTotalHours() {
  const allText = document.body.innerText;
  const match = allText.match(/Total Hours:\s*(\d+\.?\d*)\s*\(Billable:\s*(\d+\.?\d*)\)/);
  if (match) return { total: parseFloat(match[1]), billable: parseFloat(match[2]) };
  return null;
}

function readFavoriteProjects() {
  const projects = [];
  const rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMList .sapMLIB');
  for (const row of rows) {
    const cells = row.querySelectorAll('.sapMListTblCell, td');
    if (cells.length >= 2) {
      const projectText = cells[0]?.textContent?.trim();
      const activityText = cells[1]?.textContent?.trim();
      if (projectText && activityText) {
        projects.push({ project: projectText, activity: activityText });
      }
    }
  }
  if (projects.length === 0) {
    const text = document.body.innerText;
    const matches = text.match(/[A-Z]{2,5}\.\w+\s*\/\s*[^\n]+/g);
    if (matches) matches.forEach((m) => projects.push({ project: m.trim(), activity: '' }));
  }
  return projects;
}

// --- OData API Helpers ---

let csrfToken = null;

async function fetchCsrfToken() {
  try {
    const resp = await fetch(`${SAP_ODATA_BASE}/?${SAP_CLIENT}`, {
      method: 'GET',
      headers: { 'x-csrf-token': 'Fetch', 'Accept': 'application/json' },
    });
    csrfToken = resp.headers.get('x-csrf-token') || '';
    console.log('[SAP Hours Agent] CSRF token fetched');
    return csrfToken;
  } catch (e) {
    console.log('[SAP Hours Agent] CSRF fetch failed:', e);
    return '';
  }
}

async function sapODataGet(entitySet, filter, expand) {
  let url = `${SAP_ODATA_BASE}/${entitySet}?${SAP_CLIENT}&$format=json`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (expand) url += `&$expand=${expand}`;

  console.log(`[SAP Hours Agent] OData GET: ${url}`);
  const resp = await fetch(url, {
    headers: {
      'Accept': 'application/json',
      'DataServiceVersion': '2.0',
      'MaxDataServiceVersion': '2.0',
    },
  });
  if (!resp.ok) {
    const errBody = await resp.text().catch(() => '');
    console.error(`[SAP Hours Agent] OData GET failed ${resp.status}:`, errBody.substring(0, 500));
    throw new Error(`OData GET ${resp.status}: ${resp.statusText}`);
  }
  const data = await resp.json();
  console.log(`[SAP Hours Agent] OData GET ${entitySet}: ${(data.d?.results || []).length} results`);
  return data.d?.results || [];
}

async function sapODataBatchPost(entitySet, payload) {
  if (!csrfToken) await fetchCsrfToken();

  const batchBoundary = `batch_${Date.now()}`;
  const changesetBoundary = `changeset_${Date.now()}`;
  const contentId = `id-${Date.now()}`;

  const batchBody = [
    `--${batchBoundary}`,
    `Content-Type: multipart/mixed; boundary=${changesetBoundary}`,
    '',
    `--${changesetBoundary}`,
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    '',
    `POST ${entitySet}?${SAP_CLIENT} HTTP/1.1`,
    'sap-contextid-accept: header',
    'Accept: application/json',
    'Accept-Language: en',
    'DataServiceVersion: 2.0',
    'MaxDataServiceVersion: 2.0',
    'X-Requested-With: XMLHttpRequest',
    `x-csrf-token: ${csrfToken}`,
    'Content-Type: application/json',
    `Content-ID: ${contentId}`,
    `Content-Length: ${JSON.stringify(payload).length}`,
    '',
    JSON.stringify(payload),
    `--${changesetBoundary}--`,
    '',
    `--${batchBoundary}--`,
    '',
  ].join('\r\n');

  const resp = await fetch(`${SAP_ODATA_BASE}/$batch?${SAP_CLIENT}`, {
    method: 'POST',
    headers: {
      'Content-Type': `multipart/mixed;boundary=${batchBoundary}`,
      'x-csrf-token': csrfToken,
      'Accept': 'multipart/mixed',
    },
    body: batchBody,
  });

  if (!resp.ok) {
    if (resp.status === 403) {
      await fetchCsrfToken();
      return sapODataBatchPost(entitySet, payload);
    }
    throw new Error(`Batch POST ${resp.status}: ${resp.statusText}`);
  }

  const responseText = await resp.text();
  console.log('[SAP Hours Agent] Batch response FULL:', responseText);

  // Parse error from SAP batch response - look for error message in various formats
  if (responseText.includes('"error"') || responseText.match(/HTTP\/1\.1\s+[45]\d\d/)) {
    // Try structured error: {"error":{"message":{"value":"..."}}}
    const structuredErr = responseText.match(/"message"\s*:\s*\{\s*"lang"\s*:\s*"[^"]*"\s*,\s*"value"\s*:\s*"([^"]+)"/);
    // Try simple error: {"message":"..."}
    const simpleErr = responseText.match(/"message"\s*:\s*"([^"]+)"/);
    // Try innerError details
    const innerErr = responseText.match(/"innererror"[\s\S]*?"message"\s*:\s*"([^"]+)"/);

    const errMsg = structuredErr?.[1] || innerErr?.[1] || simpleErr?.[1] || 'Unknown SAP error';
    console.error('[SAP Hours Agent] Batch POST error:', errMsg);
    console.error('[SAP Hours Agent] Payload was:', JSON.stringify(payload).substring(0, 500));
    throw new Error(errMsg);
  }

  if (responseText.includes('201 Created') || responseText.includes('200 OK')) {
    const msgMatch = responseText.match(/"message":"([^"]+)"/);
    return { success: true, message: msgMatch ? msgMatch[1] : 'Entry created successfully' };
  }

  console.warn('[SAP Hours Agent] Unexpected batch response - assuming success');
  return { success: true, message: 'Entry submitted (unconfirmed)' };
}

// --- Project Search ---

async function searchProjects(query) {
  if (!query || query.length < 2) return { projects: [] };
  const q = query.toLowerCase();
  const filter = `substringof('${q}',tolower(Project)) or substringof('${q}',tolower(Description))`;
  try {
    const results = await sapODataGet('ProjectSearchSet', filter);
    console.log(`[SAP Hours Agent] Search "${query}" returned ${results.length} results`);
    return {
      projects: results.map((r) => ({ id: r.Project, name: r.Description?.trim() || '' })),
    };
  } catch (e) {
    console.error('[SAP Hours Agent] Search failed:', e);
    return { projects: [], error: e.message };
  }
}

// --- Project Activities ---

async function getProjectActivities(projectId, role) {
  if (!projectId) return { activities: [] };
  role = role || 'ZADMIN';
  const filter = `Pspid eq '${projectId}' and Role eq '${role}'`;
  try {
    const results = await sapODataGet('ProjectActivitySet', filter);
    console.log(`[SAP Hours Agent] Activities for ${projectId}: ${results.length} results`);
    return {
      activities: results.map((r) => ({
        code: r.Activity,
        description: r.Description?.trim() || '',
        projectId: r.Pspid,
      })),
    };
  } catch (e) {
    console.error('[SAP Hours Agent] getProjectActivities failed:', e);
    return { activities: [], error: e.message };
  }
}

// --- Get Recorded Hours ---

function sapDateToISO(sapDate) {
  const match = sapDate.match(/\/Date\((\d+)\)\//);
  if (!match) return sapDate;
  const d = new Date(parseInt(match[1]));
  return d.toISOString().split('T')[0];
}

async function getRecordedHours(startDate, endDate) {
  const startDt = `datetime'${startDate}T00:00:00'`;
  const endDt = `datetime'${endDate}T23:59:59'`;
  const filter = `Workdate ge ${startDt} and Workdate le ${endDt} and PersNumber eq '${PERS_NUMBER}'`;

  console.log(`[SAP Hours Agent] Getting recorded hours: ${startDate} to ${endDate}`);

  try {
    const results = await sapODataGet('TimeEntrySet', filter);
    const entries = results.map((r) => ({
      date: sapDateToISO(r.Workdate),
      project: r.Pspid,
      projectName: r.Pdesc,
      activity: r.Activity,
      activityName: r.Actdesc,
      hours: parseFloat(r.Catshours),
      status: r.Statustext || r.Status,
      description: r.Description || '',
      billing: r.Zbilling,
    }));

    const totalHours = entries.reduce((sum, e) => sum + e.hours, 0);
    const byDate = {};
    entries.forEach((e) => {
      if (!byDate[e.date]) byDate[e.date] = [];
      byDate[e.date].push(e);
    });

    return { entries, totalHours, byDate, count: entries.length };
  } catch (e) {
    return { error: e.message, entries: [], totalHours: 0 };
  }
}

// --- Get Entry Defaults ---

async function getEntryDefaults(projectId, activityCode) {
  const filter = `Pspid eq '${projectId}' and PersNumber eq '${PERS_NUMBER}' and Activity eq '${activityCode || ''}' and AbsAttType eq '0800'`;
  console.log(`[SAP Hours Agent] Getting defaults: ${projectId} / ${activityCode}`);
  try {
    const results = await sapODataGet('TimeEntryDetailSet', filter, 'NavSubSet');
    if (results.length === 0) {
      console.warn(`[SAP Hours Agent] No defaults found for ${projectId}/${activityCode}`);
      return null;
    }
    console.log(`[SAP Hours Agent] Got defaults for ${projectId}:`, JSON.stringify(results[0]).substring(0, 300));
    return results[0];
  } catch (e) {
    console.error(`[SAP Hours Agent] getEntryDefaults failed:`, e);
    return null;
  }
}

// --- Execute Actions ---

async function executeAction(action) {
  try {
    if (action.type === 'ENTER_DAY') {
      return await enterDayViaAPI(action.date, action.entries);
    }
    return { success: false, error: `Unknown action: ${action.type}` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// --- Submit time entries via OData API ---

async function enterDayViaAPI(dateStr, entries) {
  console.log(`[SAP Hours Agent] enterDayViaAPI: ${dateStr}`, JSON.stringify(entries));
  const results = [];

  for (const entry of entries) {
    try {
      let activityCode = entry.activityCode || entry.activity || '0010';
      if (activityCode && !/^\d+$/.test(activityCode)) {
        console.log(`[SAP Hours Agent] Activity "${activityCode}" looks like a name, looking up code...`);
        const actResult = await getProjectActivities(entry.projectId, 'ZADMIN');
        const match = actResult.activities.find(a =>
          a.description.toLowerCase().includes(activityCode.toLowerCase()) ||
          a.code.toLowerCase() === activityCode.toLowerCase()
        );
        if (match) {
          console.log(`[SAP Hours Agent] Resolved "${activityCode}" -> "${match.code}"`);
          activityCode = match.code;
        } else {
          console.log(`[SAP Hours Agent] Could not resolve "${activityCode}", using "0010"`);
          activityCode = '0010';
        }
      }

      const defaults = await getEntryDefaults(entry.projectId, activityCode);

      if (!defaults) {
        results.push({
          success: false,
          project: entry.projectId,
          error: `Could not get defaults for ${entry.projectId}`,
        });
        continue;
      }

      const workdate = `${dateStr}T11:00:00`;

      const payload = {
        Counter: '1',
        Workdate: workdate,
        PersNumber: PERS_NUMBER,
        UserName: USER_NAME,
        Acttype: defaults.Acttype || 'ZADMIN',
        Adesc: defaults.Adesc || 'Administrative Staff',
        Pspid: entry.projectId,
        SubProj: entry.subProject || defaults.Subproj || '',
        Activity: activityCode || defaults.Activity || '0010',
        Actdesc: defaults.Actdesc || '',
        Rework: defaults.Rework || '',
        AbsAttType: '0800',
        Oldproj: defaults.Oldproj || '',
        Billingcode: defaults.Billingcode || '',
        Zbilling: defaults.Zbilling || 'N',
        Wtart: defaults.Wtart || 'CANA',
        Zlstar: defaults.Zlstar || 'ZADMIN',
        Extref: defaults.Extref || '',
        Catshours: String(entry.hours),
        Catstime: 'P00DT00H00M00S',
        Starttime: 'PT00H00M00S',
        Endtime: 'PT00H00M00S',
        Description: entry.description || entry.projectName || 'Time entry',
        SendCctr: defaults.SendCctr || '',
        Pdesc: defaults.Pdesc || entry.projectName || '',
        Network: defaults.Network || '',
        Oldpdesc: defaults.Oldpdesc || '',
        Atext: defaults.Atext || '',
        ZbillingOld: defaults.ZbillingOld || '',
        Zbillingtype: defaults.Zbillingtype || '',
        Zvbeln: defaults.Zvbeln || '',
        Zadesc: defaults.Zadesc || '',
        Unit: defaults.Unit || '',
        AllDayFlag: false,
        Shorttext: '',
        Longtext: false,
        Row: 0,
        Status: '',
        Statustext: '',
        Statusicon: '',
        Readonly: false,
        Buttonsdisabled: false,
        Actbasedbilling: defaults.Actbasedbilling || false,
        AdminProj: defaults.AdminProj || false,
        CompanyCode: defaults.CompanyCode || '1100',
        ControllingArea: defaults.ControllingArea || '1000',
        Reason: '',
        Weekly: '',
        Clktim: defaults.Clktim || '',
        Showact: defaults.Showact || '',
        Tothrs: '',
        Billfl: defaults.Billfl || '',
        Ltreqd: defaults.Ltreqd || '',
        Ltcaps: defaults.Ltcaps || '',
        Asort: defaults.Asort || '',
        Billst: defaults.Billst || '',
        Attgroup: defaults.Attgroup || 'GL',
        Rwflag: defaults.Rwflag || '',
        Spflag: defaults.Spflag || '',
        Cotype: defaults.Cotype || '',
        Coday: defaults.Coday || '',
        Cotar: defaults.Cotar || '',
        Frdat: defaults.Frdat || null,
        Grace: defaults.Grace || '30',
        Wkhrs: defaults.Wkhrs || '00',
      };

      console.log('[SAP Hours Agent] Submitting payload:', JSON.stringify(payload, null, 2));
      const result = await sapODataBatchPost('TimeEntrySet', payload);
      results.push({
        success: true,
        project: entry.projectId,
        hours: entry.hours,
        message: result.message,
      });
    } catch (err) {
      results.push({
        success: false,
        project: entry.projectId,
        error: err.message,
      });
    }
  }

  const successCount = results.filter((r) => r.success).length;
  const failCount = results.filter((r) => !r.success).length;

  if (successCount > 0) {
    // Notify iframe to save state before reload
    const iframe = document.getElementById('sap-agent-iframe');
    if (iframe && iframe.contentWindow) {
      iframe.contentWindow.postMessage({ source: 'sap-hours-agent-response', type: 'SAVE_BEFORE_RELOAD' }, '*');
    }
    setTimeout(() => window.location.reload(), 2000);
  }

  return {
    success: failCount === 0,
    message: `${successCount}/${entries.length} entries created${failCount > 0 ? ` (${failCount} failed)` : ''}`,
    details: results,
  };
}

// --- Helpers ---

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// --- Auto-inject panel on SAP pages ---

injectAgentPanel();
