const { app } = require('@azure/functions');
const https = require('https');
const url = require('url');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

// --- Configuration (from environment or defaults) ---

const TENANT_ID = process.env.ENTRA_TENANT_ID || '2a9f86a9-29e7-44bd-8863-849373d53db8';
const CLIENT_ID = process.env.ENTRA_CLIENT_ID || '18e3ec56-ab30-427f-84cf-b3ee61e4887d';
const ISSUER = `https://login.microsoftonline.com/${TENANT_ID}/v2.0`;
const JWKS_URI = `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`;
const ALLOWED_DOMAIN = process.env.ALLOWED_EMAIL_DOMAIN || '@epiqglobal.com';

// CORS: restrict to known origins only
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '').split(',').filter(Boolean);
// Always allow the extension origin (chrome-extension:// URLs don't send Origin on POST,
// but we add the SAP hostname for iframe/fetch calls)
const SAP_HOSTNAME = process.env.SAP_HOSTNAME || 'saphub.epiqglobal.com';
if (!ALLOWED_ORIGINS.includes(`https://${SAP_HOSTNAME}`)) {
  ALLOWED_ORIGINS.push(`https://${SAP_HOSTNAME}`);
}

// Rate limiting: per-user, in-memory (resets on function app restart)
const RATE_LIMIT_WINDOW_MS = 60 * 1000; // 1 minute
const RATE_LIMIT_MAX_REQUESTS = parseInt(process.env.RATE_LIMIT_MAX || '20', 10);
const rateLimitMap = new Map();

function checkRateLimit(userId) {
  const now = Date.now();
  let entry = rateLimitMap.get(userId);
  if (!entry || now - entry.windowStart > RATE_LIMIT_WINDOW_MS) {
    entry = { windowStart: now, count: 0 };
    rateLimitMap.set(userId, entry);
  }
  entry.count++;
  return entry.count <= RATE_LIMIT_MAX_REQUESTS;
}

// Periodic cleanup of stale rate limit entries (every 5 min)
setInterval(() => {
  const cutoff = Date.now() - RATE_LIMIT_WINDOW_MS * 2;
  for (const [key, val] of rateLimitMap) {
    if (val.windowStart < cutoff) rateLimitMap.delete(key);
  }
}, 5 * 60 * 1000);

// --- JWT Verification ---

const client = jwksClient({ jwksUri: JWKS_URI, cache: true, rateLimit: true });

function getSigningKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

function verifyToken(token) {
  return new Promise((resolve, reject) => {
    jwt.verify(token, getSigningKey, {
      audience: CLIENT_ID,
      issuer: ISSUER,
      algorithms: ['RS256'],
    }, (err, decoded) => {
      if (err) return reject(err);
      resolve(decoded);
    });
  });
}

// --- CORS Helper ---

function getCorsHeaders(requestOrigin) {
  const headers = {
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Max-Age': '86400',
  };
  // Only reflect origin if it's in the allow list
  if (requestOrigin && ALLOWED_ORIGINS.includes(requestOrigin)) {
    headers['Access-Control-Allow-Origin'] = requestOrigin;
    headers['Vary'] = 'Origin';
  }
  // Chrome extensions send Origin as chrome-extension://<id> — allow those
  if (requestOrigin && requestOrigin.startsWith('chrome-extension://')) {
    headers['Access-Control-Allow-Origin'] = requestOrigin;
    headers['Vary'] = 'Origin';
  }
  return headers;
}

// --- System Prompt (server-side only — never sent to or from client) ---

function buildSystemPrompt(ctx) {
  const today = new Date().toISOString().split('T')[0];
  const dayOfWeek = new Date().toLocaleDateString('en-US', { weekday: 'long' });

  const displayName = ctx.displayName || 'User';
  const company = ctx.company || 'Unknown';
  const persNumber = ctx.persNumber || 'Unknown';
  const costCenterDisplay = ctx.costCenterName
    ? `${ctx.costCenter} / ${ctx.costCenterName}`
    : ctx.costCenter || 'Unknown';
  const defaultRole = ctx.defaultRole || 'ZADMIN';

  let projectSection = '';
  if (ctx.favorites && ctx.favorites.length > 0) {
    projectSection = 'FAVORITE PROJECTS (from SAP):\n';
    ctx.favorites.forEach((p, i) => {
      if (p.projectDesc) {
        projectSection += `${i + 1}. ${p.project} — ${p.projectDesc} (Activity: ${p.activity} / ${p.activityDesc || ''})\n`;
      } else if (typeof p === 'object') {
        projectSection += `${i + 1}. ${p.project} (Activity: ${p.activity})\n`;
      } else {
        projectSection += `${i + 1}. ${p}\n`;
      }
    });
  }

  let sapStateSection = '';
  if (ctx.sapState) {
    sapStateSection = `CURRENT SAP PAGE STATE:
Day tabs: ${JSON.stringify(ctx.sapState.dayTabs || [])}
Current entries: ${JSON.stringify(ctx.sapState.currentEntries || [])}
Total hours: ${JSON.stringify(ctx.sapState.totalHours)}`;
  }

  return `You are an AI assistant embedded in a browser extension on the SAP Fiori Time Entry page. You help ${displayName} enter their weekly hours.

SCOPE: You ONLY help with SAP time entry tasks — logging hours, checking recorded time, managing favorites, copying weeks, and answering questions about the user's timesheet. If the user asks for anything unrelated (jokes, trivia, general questions, creative writing, etc.), politely decline and redirect: "I'm your SAP time entry assistant — I can only help with logging and managing your hours. What would you like to enter?"

CONTEXT:
- Today is ${today} (${dayOfWeek})
- ${displayName} works at ${company}, Personnel #${persNumber}
- Cost Center: ${costCenterDisplay}
- Default Role: ${defaultRole}
- Standard work day: 8 hours, Mon-Fri
- Last week = Mon-Fri of the week before today's week

${projectSection || 'No favorites loaded from page.\n'}
PROJECT RESOLUTION (follow this order):
1. **Favorites first** — ALWAYS check the FAVORITE PROJECTS list above before anything else. Match by name, keyword, or partial match (e.g. "agentics" matches "Agentics AI - 2026", "admin" matches "LS DEV ADMIN"). Favorites already have the correct projectId and activity code — use them directly.
2. **Recent history** — If no favorite matches, check recent time entries with GET_RECORDED_HOURS to find projects the user has used before.
3. **Search** — Only if neither favorites nor history match, search the full project database:

\`\`\`json
{"action": "SEARCH_PROJECTS", "query": "conference"}
\`\`\`

After finding a project (from any source), you can look up its available activities:

\`\`\`json
{"action": "GET_ACTIVITIES", "projectId": "ADM.000022", "role": "ZADMIN"}
\`\`\`

LOOK UP RECORDED HOURS:
You can look up what hours have already been recorded for any date range:

\`\`\`json
{"action": "GET_RECORDED_HOURS", "startDate": "2026-02-10", "endDate": "2026-02-14"}
\`\`\`

This returns all time entries with project, activity, hours, counter (ID), status (Approved/Pending), and description. Use this when:
- The user asks what they recorded on a specific date or week
- The user wants to check if hours are already entered before adding more
- The user asks about their time entry history or status

CHECK CALENDAR STATUS:
Get a quick overview of which days in a month have hours entered:

\`\`\`json
{"action": "GET_CALENDAR_STATUS", "referenceDate": "2026-03-15"}
\`\`\`

Returns each day that has hours, whether it's complete (8h), and gaps. Use this to quickly identify missing days without fetching full entry details.

GET WEEK TOTAL:
Get the current week's total hours in a single call:

\`\`\`json
{"action": "GET_WEEK_TOTAL"}
\`\`\`

DELETE AN ENTRY:
Delete a specific time entry by its counter ID (from GET_RECORDED_HOURS results):

\`\`\`json
{"action": "DELETE_ENTRY", "counter": "000055693654"}
\`\`\`

Always confirm with the user before deleting. Show them what will be deleted first.

COPY WEEK:
Copy all entries from one week's Monday to another week's Monday:

\`\`\`json
{"action": "COPY_WEEK", "fromDate": "2026-03-02", "toDate": "2026-03-09", "move": false}
\`\`\`

Set move=true to move entries instead of copy. Confirm with user first.

MANAGE FAVORITES:
Add or remove projects from the user's SAP favorites:

\`\`\`json
{"action": "ADD_FAVORITE", "projectId": "DEV.000982", "activity": "0010", "description": ""}
\`\`\`

\`\`\`json
{"action": "REMOVE_FAVORITE", "projectId": "DEV.000982", "activity": "0010"}
\`\`\`

${sapStateSection}

WORKFLOW:
1. User describes their week in natural language
2. Resolve projects using PROJECT RESOLUTION order above (favorites → history → search)
3. Infer reasonable defaults instead of asking:
   - "each day" or "all week" = Monday–Friday of the current or most recent work week
   - Activity code: use the one from the matched favorite, or the most common one from history
   - Description: generate a professional description from what the user said + the project name
   - Only ask if there is genuine ambiguity (e.g. multiple projects could match, or date range is unclear)
4. Propose the time entries with a clear summary table
5. On user confirmation, submit with ENTER_TIME

When ready to submit, include a JSON block with this exact format:

\`\`\`json
{"action": "ENTER_TIME", "entries": [
  {"date": "YYYY-MM-DD", "projectId": "DEV.000982", "projectName": "Agentics AI - 2026", "activity": "Coding", "hours": 8, "description": "Worked on Agentics AI coding"},
  ...
]}
\`\`\`

IMPORTANT: Every entry MUST include a "description" field (SAP rejects entries without comments). Use a brief, professional description of the work.

RULES:
- Always confirm the plan with the user BEFORE including the ENTER_TIME, DELETE_ENTRY, or COPY_WEEK actions
- When proposing entries, ALWAYS show a clear table/list that includes the date, project, hours, AND the description you plan to use for each entry. The description is a required field in SAP and will be visible to managers, so the user must verify it.
- If the user says "yes", "do it", "go ahead", "submit", "looks good", etc., THEN include the action JSON
- NEVER use SEARCH_PROJECTS for projects that match a favorite — use the favorite's projectId and activity directly
- Use GET_CALENDAR_STATUS to quickly check which days need hours before proposing entries
- Days should total 8 hours unless the user says otherwise
- If the user says they took a day off, skip that day entirely (PTO is handled separately in SAP)
- Generate professional, accurate descriptions based on what the user told you (e.g. "Agentics AI development and coding", "AI Platform requirements and design")
- Be concise — propose entries on your first response when you have enough info. Don't ask clarifying questions you can infer the answer to.
- Only ask when there is genuine ambiguity that could lead to wrong entries`;
}

// --- Input Validation ---

const MAX_MESSAGE_LENGTH = 2000;   // per individual message content
const MAX_MESSAGES = 50;           // conversation history limit
const MAX_CONTEXT_FIELD = 500;     // per context string field
const ALLOWED_ROLES = ['user', 'assistant'];

function validateMessages(messages) {
  if (!Array.isArray(messages)) return 'messages must be an array';
  if (messages.length > MAX_MESSAGES) return `Too many messages (max ${MAX_MESSAGES})`;

  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    if (!msg || typeof msg !== 'object') return `messages[${i}] must be an object`;
    if (!ALLOWED_ROLES.includes(msg.role)) return `messages[${i}].role must be "user" or "assistant"`;
    if (typeof msg.content !== 'string') return `messages[${i}].content must be a string`;
    if (msg.content.length > MAX_MESSAGE_LENGTH) return `messages[${i}].content too long (max ${MAX_MESSAGE_LENGTH} chars)`;
  }
  return null;
}

function validateContext(ctx) {
  if (!ctx || typeof ctx !== 'object') return 'context must be an object';

  // Validate string fields don't exceed limits
  const stringFields = ['displayName', 'company', 'persNumber', 'costCenter', 'costCenterName', 'defaultRole'];
  for (const field of stringFields) {
    if (ctx[field] !== undefined) {
      if (typeof ctx[field] !== 'string') return `context.${field} must be a string`;
      if (ctx[field].length > MAX_CONTEXT_FIELD) return `context.${field} too long`;
    }
  }

  // Validate favorites array
  if (ctx.favorites !== undefined) {
    if (!Array.isArray(ctx.favorites)) return 'context.favorites must be an array';
    if (ctx.favorites.length > 50) return 'Too many favorites';
  }

  return null;
}

// --- Sanitize context strings (prevent prompt injection via user profile fields) ---

function sanitizeString(str, maxLen) {
  if (typeof str !== 'string') return '';
  // Remove control characters and truncate
  return str.replace(/[\x00-\x1f\x7f]/g, '').slice(0, maxLen || MAX_CONTEXT_FIELD);
}

function sanitizeContext(ctx) {
  return {
    displayName: sanitizeString(ctx.displayName, 100),
    company: sanitizeString(ctx.company, 100),
    persNumber: sanitizeString(ctx.persNumber, 20),
    costCenter: sanitizeString(ctx.costCenter, 20),
    costCenterName: sanitizeString(ctx.costCenterName, 100),
    defaultRole: sanitizeString(ctx.defaultRole, 20),
    favorites: (ctx.favorites || []).slice(0, 50).map(f => ({
      project: sanitizeString(f.project, 50),
      projectDesc: sanitizeString(f.projectDesc, 200),
      activity: sanitizeString(f.activity, 20),
      activityDesc: sanitizeString(f.activityDesc, 200),
      network: sanitizeString(f.network, 50),
    })),
    sapState: ctx.sapState ? {
      dayTabs: (ctx.sapState.dayTabs || []).slice(0, 10),
      currentEntries: (ctx.sapState.currentEntries || []).slice(0, 50),
      totalHours: ctx.sapState.totalHours || null,
    } : null,
  };
}

// --- Azure Function Handler ---

app.http('claude', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const origin = request.headers.get('origin') || '';
    const corsHeaders = getCorsHeaders(origin);

    // CORS preflight
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: corsHeaders };
    }

    // Validate Bearer token
    const authHeader = request.headers.get('authorization') || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    if (!token) {
      return { status: 401, headers: corsHeaders, jsonBody: { error: 'Authorization required' } };
    }

    let claims;
    try {
      claims = await verifyToken(token);
    } catch (e) {
      context.error('Token validation failed:', e.message);
      // Don't leak JWT error details to client
      return { status: 401, headers: corsHeaders, jsonBody: { error: 'Authentication failed' } };
    }

    // Restrict to allowed domain
    const email = (claims.preferred_username || claims.email || '').toLowerCase();
    if (!email.endsWith(ALLOWED_DOMAIN)) {
      return { status: 403, headers: corsHeaders, jsonBody: { error: 'Access restricted' } };
    }

    // Rate limiting
    const userId = claims.oid || claims.sub || email;
    if (!checkRateLimit(userId)) {
      context.warn(`Rate limit exceeded for ${email}`);
      return { status: 429, headers: corsHeaders, jsonBody: { error: 'Too many requests. Please wait a moment.' } };
    }

    context.log(`Request from: ${email}`);

    // Parse and validate request body
    let body;
    try {
      body = await request.json();
    } catch (e) {
      return { status: 400, headers: corsHeaders, jsonBody: { error: 'Invalid JSON' } };
    }

    if (!body || !body.messages) {
      return { status: 400, headers: corsHeaders, jsonBody: { error: 'Request must include messages array' } };
    }

    // Validate messages
    const msgError = validateMessages(body.messages);
    if (msgError) {
      return { status: 400, headers: corsHeaders, jsonBody: { error: msgError } };
    }

    // Validate and sanitize context
    const ctxError = validateContext(body.context || {});
    if (ctxError) {
      return { status: 400, headers: corsHeaders, jsonBody: { error: ctxError } };
    }

    const sanitizedCtx = sanitizeContext(body.context || {});

    // Build system prompt SERVER-SIDE (client cannot control this)
    const systemPrompt = buildSystemPrompt(sanitizedCtx);

    // Server-controlled Claude API parameters
    const apiKey = process.env.AZURE_CLAUDE_API_KEY;
    const endpoint = process.env.AZURE_CLAUDE_ENDPOINT;
    const model = process.env.AZURE_CLAUDE_MODEL || 'claude-sonnet-4-5';
    const apiVersion = process.env.AZURE_CLAUDE_API_VERSION || '2023-06-01';
    const maxTokens = parseInt(process.env.MAX_TOKENS || '1024', 10);

    if (!apiKey || !endpoint) {
      context.error('Missing AZURE_CLAUDE_API_KEY or AZURE_CLAUDE_ENDPOINT');
      return { status: 500, headers: corsHeaders, jsonBody: { error: 'Service misconfigured' } };
    }

    // Build Claude request — ALL parameters server-controlled
    const payload = {
      model: model,
      max_tokens: maxTokens,
      system: systemPrompt,
      messages: body.messages,
    };

    try {
      const parsed = new url.URL(endpoint);
      const postData = JSON.stringify(payload);

      const result = await new Promise((resolve, reject) => {
        const options = {
          hostname: parsed.hostname,
          port: 443,
          path: `${parsed.pathname}?api-version=${apiVersion}`,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': apiVersion,
            'Content-Length': Buffer.byteLength(postData),
          },
        };

        const req = https.request(options, (response) => {
          const chunks = [];
          response.on('data', (chunk) => chunks.push(chunk));
          response.on('end', () => {
            resolve({ statusCode: response.statusCode, body: Buffer.concat(chunks).toString() });
          });
        });

        req.on('error', reject);
        req.setTimeout(120000, () => {
          req.destroy();
          reject(new Error('Request to Claude API timed out'));
        });
        req.write(postData);
        req.end();
      });

      return {
        status: result.statusCode,
        headers: { 'Content-Type': 'application/json', ...corsHeaders },
        body: result.body,
      };
    } catch (e) {
      context.error('Claude API error:', e.message);
      return { status: 502, headers: corsHeaders, jsonBody: { error: 'Failed to reach AI service' } };
    }
  },
});
