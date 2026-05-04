// SAP Fiori Time Entry - Content Script
// CONFIG is loaded before this script via manifest.json content_scripts order

console.log('[SAP Hours Agent] Content script loaded on:', window.location.pathname);

// Skip non-FLP pages (SAML auth callbacks, etc.) — discovery and panel are useless there
const NON_FLP_PATHS = ['/sap/saml2/', '/sap/public/', '/sap/bc/sec/'];
const isNonFLPPage = NON_FLP_PATHS.some(p => window.location.pathname.startsWith(p));
if (isNonFLPPage) {
  console.log('[SAP Hours Agent] Skipping — not an FLP page');
}

const SAP_ODATA_BASE = CONFIG.sapODataPath;
const SAP_CLIENT = `sap-client=${CONFIG.sapClient}`;

// Sanitize values interpolated into OData $filter strings to prevent injection.
// OData string literals use single quotes; escape by doubling them.
// Also strip control characters that could break the query.
function odataSanitize(val) {
  if (typeof val !== 'string') return '';
  return val.replace(/[\x00-\x1f\x7f]/g, '').replace(/'/g, "''");
}

// Extension origin for secure postMessage targeting
const EXTENSION_ORIGIN = chrome.runtime.getURL('').replace(/\/$/, '');

// User info - populated by auto-discovery
let PERS_NUMBER = '';
let USER_NAME = '';
let userConfig = null;
let iframeReady = false;

// --- Auto-Discovery: detect user info from SAP APIs ---

async function discoverUserInfo() {
  // Check cache first
  const cached = await new Promise(resolve => {
    chrome.storage.local.get('sapUserConfig', data => resolve(data.sapUserConfig));
  });

  if (cached && cached.persNumber && Date.now() - cached.discoveredAt < 7 * 24 * 60 * 60 * 1000) {
    console.log('[SAP Hours Agent] Using cached user config:', cached.displayName);
    return cached;
  }

  console.log('[SAP Hours Agent] Auto-discovering user info from SAP...');

  const config = {
    persNumber: '',
    userName: '',
    displayName: '',
    userTitle: '',
    company: '',
    costCenter: '',
    costCenterName: '',
    defaultRole: 'ZADMIN',
    location: '',
    workHours: '08',
    sapHostname: CONFIG.sapHostname,
    discoveredAt: Date.now(),
  };

  // 1. Read username from the FLP server-rendered meta tag
  //    SAP embeds startupConfig (with user id, email, name) into a <meta> tag
  try {
    const metaEls = document.querySelectorAll('meta[name^="sap.ushellConfig.serverSideConfig"]');
    for (const meta of metaEls) {
      const content = meta.getAttribute('content');
      if (content && content.includes('startupConfig')) {
        const parsed = JSON.parse(content);
        const startup = parsed.startupConfig || parsed;
        config.userName = startup.id || '';
        config.displayName = startup.fullName || '';
        console.log(`[SAP Hours Agent] FLP meta tag: user=${config.userName}, name=${config.displayName}`);
        break;
      }
    }
  } catch (e) {
    console.log('[SAP Hours Agent] FLP meta tag parse failed:', e.message);
  }

  // 2. Query HRInfoSet for full profile (personnel number, cost center, etc.)
  //    Field names verified from HAR: PersNumber, FulllName (3 L's), CompName,
  //    CostCenter, CostCenterName, ActivityType, JobTitle, Location, WorkHours
  if (config.userName) {
    try {
      const filter = `UserName eq '${config.userName}'`;
      const results = await sapODataGet('HRInfoSet', filter);
      if (results.length > 0) {
        const hr = results[0];
        config.persNumber = hr.PersNumber || '';
        config.displayName = config.displayName || hr.FulllName || '';
        config.company = hr.CompName || '';
        config.costCenter = hr.CostCenter || '';
        config.costCenterName = hr.CostCenterName || '';
        config.defaultRole = hr.ActivityType || 'ZADMIN';
        config.userTitle = hr.JobTitle || '';
        config.location = hr.Location || '';
        config.workHours = hr.WorkHours || '08';
        console.log(`[SAP Hours Agent] HRInfoSet: pers=${config.persNumber}, costCenter=${config.costCenterName}`);
      }
    } catch (e) {
      console.log('[SAP Hours Agent] HRInfoSet failed:', e.message);
    }
  }

  // 3. Fallback if no username from meta tag: try HRInfoSet by key (empty = current user)
  if (!config.userName) {
    try {
      // Some SAP systems allow querying HRInfoSet without a filter for the current user
      const results = await sapODataGet('HRInfoSet', '');
      if (results.length > 0) {
        const hr = results[0];
        config.userName = hr.UserName || '';
        config.persNumber = hr.PersNumber || '';
        config.displayName = hr.FulllName || '';
        config.company = hr.CompName || '';
        config.costCenter = hr.CostCenter || '';
        config.costCenterName = hr.CostCenterName || '';
        config.defaultRole = hr.ActivityType || 'ZADMIN';
        config.userTitle = hr.JobTitle || '';
        config.location = hr.Location || '';
        config.workHours = hr.WorkHours || '08';
        console.log(`[SAP Hours Agent] HRInfoSet (no filter): pers=${config.persNumber}`);
      }
    } catch (e) {
      console.log('[SAP Hours Agent] HRInfoSet fallback failed:', e.message);
    }
  }

  // 4. Last resort: scrape the page for user info
  if (!config.persNumber || !config.displayName) {
    try {
      if (!config.displayName) {
        const shellHeader = document.querySelector('#meAreaHeaderButton .sapUShellShellHeadUsrItmName');
        if (shellHeader) config.displayName = shellHeader.textContent.trim();
      }
      console.log(`[SAP Hours Agent] Page scrape: name=${config.displayName}`);
    } catch (e) {
      console.log('[SAP Hours Agent] Page scrape failed:', e.message);
    }
  }

  // Cache result
  if (config.persNumber) {
    chrome.storage.local.set({ sapUserConfig: config });
    console.log('[SAP Hours Agent] User config discovered and cached:', config.displayName, config.persNumber);
  } else {
    console.warn('[SAP Hours Agent] Could not discover personnel number. Will retry...');
  }

  return config;
}

// Retry discovery after a delay (FLP may still be bootstrapping)
let discoveryRetries = 0;
const MAX_DISCOVERY_RETRIES = 3;
const DISCOVERY_RETRY_DELAY = 3000; // 3 seconds

async function retryDiscoveryIfNeeded() {
  if (PERS_NUMBER || discoveryRetries >= MAX_DISCOVERY_RETRIES) return;
  discoveryRetries++;
  console.log(`[SAP Hours Agent] Retry discovery attempt ${discoveryRetries}/${MAX_DISCOVERY_RETRIES}...`);
  // Clear cache so we re-discover fresh
  await new Promise(resolve => chrome.storage.local.remove('sapUserConfig', resolve));
  await initUserConfig();
  if (!PERS_NUMBER && discoveryRetries < MAX_DISCOVERY_RETRIES) {
    setTimeout(retryDiscoveryIfNeeded, DISCOVERY_RETRY_DELAY);
  }
}

// --- Initialize ---

async function initUserConfig() {
  userConfig = await discoverUserInfo();
  PERS_NUMBER = userConfig.persNumber;
  USER_NAME = userConfig.userName;

  // Notify the iframe that config is ready (only if iframe has loaded)
  const iframe = document.getElementById('sap-agent-iframe');
  if (iframe && iframe.contentWindow && iframeReady) {
    iframe.contentWindow.postMessage({
      source: 'sap-hours-agent-response',
      type: 'USER_CONFIG_READY',
      config: userConfig,
    }, EXTENSION_ORIGIN);
  }
}

// Start discovery immediately (skip on auth/SAML pages)
if (!isNonFLPPage) {
  initUserConfig().then(() => {
    if (!PERS_NUMBER) {
      setTimeout(retryDiscoveryIfNeeded, DISCOVERY_RETRY_DELAY);
    }
  });
}

// --- Inject Agent Panel into SAP page ---

let panelVisible = true;

function injectAgentPanel() {
  if (document.getElementById('sap-hours-agent-root')) return;

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

  const iframe = document.createElement('iframe');
  iframe.id = 'sap-agent-iframe';
  iframe.src = chrome.runtime.getURL('sidepanel.html');
  iframe.style.cssText = `
    width: 100%;
    height: 100%;
    border: none;
    background: #1a1a2e;
  `;
  iframe.addEventListener('load', () => {
    iframeReady = true;
    if (userConfig) {
      iframe.contentWindow.postMessage({
        source: 'sap-hours-agent-response',
        type: 'USER_CONFIG_READY',
        config: userConfig,
      }, EXTENSION_ORIGIN);
    }
  });

  panel.appendChild(iframe);
  root.appendChild(toggle);
  root.appendChild(panel);
  document.body.appendChild(root);

  toggle.addEventListener('click', () => {
    togglePanel();
  });

  console.log('[SAP Hours Agent] Panel injected');
}

function togglePanel() {
  const panel = document.getElementById('sap-agent-panel');
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
  if (msg.type === 'GET_USER_CONFIG') {
    const respond = async () => {
      if (!userConfig) await initUserConfig();
      sendResponse(userConfig);
    };
    respond();
    return true;
  }
  if (msg.type === 'GET_FAVORITES') {
    getFavoriteProjects().then(sendResponse).catch(e => sendResponse([]));
    return true;
  }
  if (msg.type === 'GET_CALENDAR_STATUS') {
    getCalendarStatus(msg.referenceDate).then(sendResponse).catch(e => sendResponse({ days: [], error: e.message }));
    return true;
  }
  if (msg.type === 'GET_WEEK_TOTAL') {
    getWeekTotal().then(sendResponse).catch(e => sendResponse({ hours: 0, error: e.message }));
    return true;
  }
});

// --- PostMessage Bridge (from embedded iframe) ---

window.addEventListener('message', async (event) => {
  if (!event.data || event.data.source !== 'sap-hours-agent') return;
  // Only accept messages from our extension iframe
  if (event.origin !== EXTENSION_ORIGIN) return;

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
      case 'RELOAD_PAGE':
        setTimeout(() => window.location.reload(), 1000);
        result = { ok: true };
        break;
      case 'NAVIGATE_TIME_ENTRY':
        setTimeout(() => {
          window.location.href = `https://${CONFIG.sapHostname}/sap/bc/ui2/flp#FioriTime-Enter`;
        }, 300);
        result = { ok: true };
        break;
      case 'GET_USER_CONFIG':
        if (!userConfig) await initUserConfig();
        result = userConfig;
        break;
      case 'GET_FAVORITES':
        result = await getFavoriteProjects();
        break;
      case 'GET_CALENDAR_STATUS':
        result = await getCalendarStatus(payload.referenceDate);
        break;
      case 'GET_WEEK_TOTAL':
        result = await getWeekTotal();
        break;
      default:
        result = { error: `Unknown message type: ${type}` };
    }
  } catch (e) {
    console.error(`[SAP Hours Agent] postMessage handler error (${type}):`, e);
    result = { error: e.message };
  }

  const iframe = document.getElementById('sap-agent-iframe');
  if (iframe && iframe.contentWindow) {
    iframe.contentWindow.postMessage({
      source: 'sap-hours-agent-response',
      id,
      result,
    }, EXTENSION_ORIGIN);
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
    const favoriteProjects = await getFavoriteProjects();

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

async function getFavoriteProjects() {
  // Get favorites via OData (NavFavSet) — much more reliable than DOM scraping
  if (!USER_NAME) return [];
  try {
    const filter = `UserName eq '${USER_NAME}'`;
    const results = await sapODataGet('HRInfoSet', filter, 'NavFavSet');
    if (results.length > 0 && results[0].NavFavSet) {
      const favs = results[0].NavFavSet.results || [];
      return favs.map(f => ({
        project: f.ProjectName,
        projectDesc: f.ProjectDesc,
        activity: f.Activity,
        activityDesc: f.ActivityDesc,
        network: f.Network,
      }));
    }
  } catch (e) {
    console.log('[SAP Hours Agent] NavFavSet fetch failed, falling back to DOM:', e.message);
  }
  // Fallback to DOM scraping
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
  return projects;
}

// --- Calendar Status (which days have hours) ---

async function getCalendarStatus(referenceDate) {
  if (!PERS_NUMBER) return { days: [], error: 'No personnel number' };
  const workHours = userConfig?.workHours || '08';
  const dateStr = `datetime'${referenceDate}T15:00:00'`;
  const filter = `PersNumber eq '${PERS_NUMBER}' and Wkhrs eq '${workHours}' and Date eq ${dateStr}`;
  try {
    const results = await sapODataGet('CalendarMarkingSet', filter);
    const days = results.map(r => {
      const isoDate = sapDateToISO(r.Date);
      const dayName = new Date(isoDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'long' });
      return {
        date: isoDate,
        dayOfWeek: dayName,
        hours: parseFloat(r.Hours),
        tooltip: r.Tooltip,
        color: r.Color,
        complete: r.Color === 'limegreen',
      };
    });
    return { days };
  } catch (e) {
    return { days: [], error: e.message };
  }
}

// --- Week Total (lightweight) ---

async function getWeekTotal() {
  try {
    const url = `${SAP_ODATA_BASE}/DynTileInfoSet('WeekHours')?${SAP_CLIENT}&$format=json&sap-language=EN`;
    const resp = await fetch(url, {
      headers: { 'Accept': 'application/json' },
    });
    if (!resp.ok) throw new Error(`${resp.status}`);
    const data = await resp.json();
    const d = data.d || data;
    return {
      hours: parseFloat(d.number) || 0,
      label: d.numberUnit || 'Hours this week',
    };
  } catch (e) {
    return { hours: 0, error: e.message };
  }
}

// --- Delete Time Entry ---

async function deleteTimeEntry(counter) {
  if (!csrfToken) await fetchCsrfToken();
  const url = `${SAP_ODATA_BASE}/TimeEntrySet('${odataSanitize(counter)}')?${SAP_CLIENT}`;
  console.log(`[SAP Hours Agent] DELETE: ${url}`);
  const resp = await fetch(url, {
    method: 'DELETE',
    headers: {
      'x-csrf-token': csrfToken,
      'X-Requested-With': 'XMLHttpRequest',
    },
  });
  if (!resp.ok) {
    if (resp.status === 403) {
      await fetchCsrfToken();
      return deleteTimeEntry(counter);
    }
    const errText = await resp.text().catch(() => '');
    throw new Error(`DELETE failed ${resp.status}: ${errText.substring(0, 200)}`);
  }
  return { success: true, message: `Entry ${counter} deleted` };
}

// --- Copy/Move Entries Between Weeks ---

async function copyWeekEntries(fromDate, toDate, move) {
  if (!csrfToken) await fetchCsrfToken();

  // First get entries from the source week
  const fromStart = fromDate;
  const fromEnd = new Date(new Date(fromDate + 'T12:00:00').getTime() + 4 * 86400000).toISOString().split('T')[0];
  const sourceEntries = await getRecordedHours(fromStart, fromEnd);

  if (!sourceEntries.entries || sourceEntries.entries.length === 0) {
    return { success: false, error: 'No entries found in source week to copy' };
  }

  const payload = {
    DateFrom: `${fromDate}T11:00:00`,
    DateTo: `${toDate}T11:00:00`,
    Move: move ? 'X' : '',
    NavTimeEntrySet: sourceEntries.entries.map(e => ({
      Counter: e.counter || '1',
      Workdate: `${e.date}T11:00:00`,
      PersNumber: PERS_NUMBER,
      Pspid: e.project,
      Activity: e.activity,
      Catshours: String(e.hours),
      AbsAttType: '0800',
    })),
  };

  const result = await sapODataBatchPost('TimeEntryCopyToSet', payload);
  return result;
}

// --- Manage Favorites ---

async function addFavorite(projectId, activityCode, description) {
  if (!csrfToken) await fetchCsrfToken();
  const payload = {
    PersNumber: PERS_NUMBER,
    ProjectName: projectId,
    Activity: activityCode,
    Description: description || '',
  };
  return sapODataBatchPost('FavoriteSet', payload);
}

async function removeFavorite(projectId, activityCode) {
  if (!csrfToken) await fetchCsrfToken();
  const url = `${SAP_ODATA_BASE}/FavoriteSet(PersNumber='${odataSanitize(PERS_NUMBER)}',ProjectName='${odataSanitize(projectId)}',Activity='${odataSanitize(activityCode)}')?${SAP_CLIENT}`;
  console.log(`[SAP Hours Agent] DELETE favorite: ${url}`);
  const resp = await fetch(url, {
    method: 'DELETE',
    headers: {
      'x-csrf-token': csrfToken,
      'X-Requested-With': 'XMLHttpRequest',
    },
  });
  if (!resp.ok) {
    if (resp.status === 403) {
      await fetchCsrfToken();
      return removeFavorite(projectId, activityCode);
    }
    throw new Error(`DELETE favorite failed ${resp.status}`);
  }
  return { success: true, message: `Favorite ${projectId}/${activityCode} removed` };
}

// --- OData API Helpers ---

let csrfToken = null;

async function fetchCsrfToken() {
  try {
    const resp = await fetch(`${SAP_ODATA_BASE}/?${SAP_CLIENT}`, {
      method: 'HEAD',
      headers: {
        'x-csrf-token': 'Fetch',
        'X-Requested-With': 'XMLHttpRequest',
      },
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
      'DataServiceVersion': '2.0',
      'MaxDataServiceVersion': '2.0',
      'X-Requested-With': 'XMLHttpRequest',
      'sap-cancel-on-close': 'true',
      'sap-contextid-accept': 'header',
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

  if (responseText.includes('"error"') || responseText.match(/HTTP\/1\.1\s+[45]\d\d/)) {
    const structuredErr = responseText.match(/"message"\s*:\s*\{\s*"lang"\s*:\s*"[^"]*"\s*,\s*"value"\s*:\s*"([^"]+)"/);
    const simpleErr = responseText.match(/"message"\s*:\s*"([^"]+)"/);
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
  const q = odataSanitize(query.toLowerCase());
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
  const filter = `Pspid eq '${odataSanitize(projectId)}' and Role eq '${odataSanitize(role)}'`;
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
  const startDt = `datetime'${odataSanitize(startDate)}T00:00:00'`;
  const endDt = `datetime'${odataSanitize(endDate)}T23:59:59'`;
  const filter = `Workdate ge ${startDt} and Workdate le ${endDt} and PersNumber eq '${odataSanitize(PERS_NUMBER)}'`;

  console.log(`[SAP Hours Agent] Getting recorded hours: ${startDate} to ${endDate}`);

  try {
    const results = await sapODataGet('TimeEntrySet', filter);
    const entries = results.map((r) => {
      const isoDate = sapDateToISO(r.Workdate);
      const dayName = new Date(isoDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'long' });
      return {
        counter: r.Counter,
        date: isoDate,
        dayOfWeek: dayName,
        project: r.Pspid,
        projectName: r.Pdesc,
        activity: r.Activity,
        activityName: r.Actdesc,
        hours: parseFloat(r.Catshours),
        status: r.Statustext || r.Status,
        description: r.Description || '',
        billing: r.Zbilling,
      };
    });

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
  const filter = `Pspid eq '${odataSanitize(projectId)}' and PersNumber eq '${odataSanitize(PERS_NUMBER)}' and Activity eq '${odataSanitize(activityCode || '')}' and AbsAttType eq '0800'`;
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
    if (action.type === 'DELETE_ENTRY') {
      return await deleteTimeEntry(action.counter);
    }
    if (action.type === 'COPY_WEEK') {
      return await copyWeekEntries(action.fromDate, action.toDate, action.move);
    }
    if (action.type === 'ADD_FAVORITE') {
      return await addFavorite(action.projectId, action.activity, action.description);
    }
    if (action.type === 'REMOVE_FAVORITE') {
      return await removeFavorite(action.projectId, action.activity);
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
        Oldproj: entry.projectId,
        Billingcode: defaults.Billingcode || '',
        Zbilling: defaults.Zbilling || 'N',
        Wtart: defaults.Wtart || 'CANA',
        Zlstar: defaults.Zlstar || 'ZADMIN',
        Extref: defaults.Extref || '',
        Catshours: String(entry.hours),
        Catstime: 'PT00H00M00S',
        Starttime: 'PT00H00M00S',
        Endtime: 'PT00H00M00S',
        Description: entry.description || entry.projectName || 'Time entry',
        SendCctr: defaults.SendCctr || '',
        Pdesc: defaults.Pdesc || entry.projectName || '',
        Network: defaults.Network || '',
        Oldpdesc: defaults.Pdesc || entry.projectName || '',
        Atext: defaults.Atext || '',
        ZbillingOld: defaults.ZbillingOld || '',
        Zbillingtype: defaults.Zbillingtype || '',
        Zvbeln: defaults.Zvbeln || '',
        Zadesc: defaults.Zadesc || '',
        Unit: defaults.Unit || '',
        AllDayFlag: false,
        Shorttext: defaults.Shorttext || 'Fiori Time Entry',
        Longtext: true,
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

  return {
    success: failCount === 0,
    successCount,
    message: `${successCount}/${entries.length} entries created${failCount > 0 ? ` (${failCount} failed)` : ''}`,
    details: results,
  };
}

// --- Auto-inject panel on SAP pages (skip auth/SAML pages) ---

if (!isNonFLPPage) {
  injectAgentPanel();
}
