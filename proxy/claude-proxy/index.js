const https = require('https');
const url = require('url');

module.exports = async function (context, req) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    context.res = {
      status: 204,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Max-Age': '86400',
      },
    };
    return;
  }

  // Entra ID auth is enforced at the App Service level.
  // The x-ms-client-principal header is injected by Easy Auth after token validation.
  const principal = req.headers['x-ms-client-principal'];
  if (!principal) {
    context.res = { status: 401, body: { error: 'Authentication required' } };
    return;
  }

  // Decode and verify the caller is from @epiqglobal.com
  try {
    const decoded = JSON.parse(Buffer.from(principal, 'base64').toString('utf-8'));
    const claims = decoded.claims || [];
    const email = (claims.find(c => c.typ === 'preferred_username') || {}).val || '';
    if (!email.toLowerCase().endsWith('@epiqglobal.com')) {
      context.res = { status: 403, body: { error: 'Access restricted to Epiq employees' } };
      return;
    }
    context.log(`Claude API request from: ${email}`);
  } catch (e) {
    context.log.error('Failed to decode principal:', e.message);
    context.res = { status: 401, body: { error: 'Invalid authentication token' } };
    return;
  }

  // Validate request body
  if (!req.body || !req.body.messages) {
    context.res = { status: 400, body: { error: 'Request body must include messages array' } };
    return;
  }

  // Forward to Claude API with server-side credentials
  const apiKey = process.env.AZURE_CLAUDE_API_KEY;
  const endpoint = process.env.AZURE_CLAUDE_ENDPOINT;
  const model = process.env.AZURE_CLAUDE_MODEL || 'claude-sonnet-4-5';
  const apiVersion = process.env.AZURE_CLAUDE_API_VERSION || '2023-06-01';

  if (!apiKey || !endpoint) {
    context.log.error('Missing AZURE_CLAUDE_API_KEY or AZURE_CLAUDE_ENDPOINT in app settings');
    context.res = { status: 500, body: { error: 'Proxy misconfigured' } };
    return;
  }

  // Build the Claude request payload — pass through allowed fields only
  const payload = {
    model: model,
    max_tokens: req.body.max_tokens || 4096,
    messages: req.body.messages,
  };
  if (req.body.system) payload.system = req.body.system;
  if (req.body.temperature !== undefined) payload.temperature = req.body.temperature;

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

      const request = https.request(options, (response) => {
        const chunks = [];
        response.on('data', (chunk) => chunks.push(chunk));
        response.on('end', () => {
          const body = Buffer.concat(chunks).toString();
          resolve({ statusCode: response.statusCode, body });
        });
      });

      request.on('error', reject);
      request.setTimeout(120000, () => {
        request.destroy();
        reject(new Error('Request to Claude API timed out'));
      });
      request.write(postData);
      request.end();
    });

    context.res = {
      status: result.statusCode,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
      body: result.body,
      isRaw: true,
    };
  } catch (e) {
    context.log.error('Claude API error:', e.message);
    context.res = {
      status: 502,
      body: { error: 'Failed to reach Claude API' },
    };
  }
};
