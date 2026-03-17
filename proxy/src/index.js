const { app } = require('@azure/functions');
const https = require('https');
const url = require('url');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const TENANT_ID = '2a9f86a9-29e7-44bd-8863-849373d53db8';
const CLIENT_ID = '18e3ec56-ab30-427f-84cf-b3ee61e4887d';
const ISSUER = `https://login.microsoftonline.com/${TENANT_ID}/v2.0`;
const JWKS_URI = `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`;

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

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Max-Age': '86400',
};

app.http('claude', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    // CORS preflight
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: CORS_HEADERS };
    }

    // Validate Bearer token
    const authHeader = request.headers.get('authorization') || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    if (!token) {
      return { status: 401, headers: CORS_HEADERS, jsonBody: { error: 'Authorization header required' } };
    }

    let claims;
    try {
      claims = await verifyToken(token);
    } catch (e) {
      context.error('Token validation failed:', e.message);
      return { status: 401, headers: CORS_HEADERS, jsonBody: { error: 'Invalid token: ' + e.message } };
    }

    // Restrict to @epiqglobal.com
    const email = (claims.preferred_username || claims.email || '').toLowerCase();
    if (!email.endsWith('@epiqglobal.com')) {
      return { status: 403, headers: CORS_HEADERS, jsonBody: { error: 'Access restricted to Epiq employees' } };
    }
    context.log(`Claude API request from: ${email}`);

    // Parse request body
    const body = await request.json();
    if (!body || !body.messages) {
      return { status: 400, headers: CORS_HEADERS, jsonBody: { error: 'Request body must include messages array' } };
    }

    // Forward to Claude API with server-side credentials
    const apiKey = process.env.AZURE_CLAUDE_API_KEY;
    const endpoint = process.env.AZURE_CLAUDE_ENDPOINT;
    const model = process.env.AZURE_CLAUDE_MODEL || 'claude-sonnet-4-5';
    const apiVersion = process.env.AZURE_CLAUDE_API_VERSION || '2023-06-01';

    if (!apiKey || !endpoint) {
      context.error('Missing AZURE_CLAUDE_API_KEY or AZURE_CLAUDE_ENDPOINT in app settings');
      return { status: 500, headers: CORS_HEADERS, jsonBody: { error: 'Proxy misconfigured' } };
    }

    // Build Claude request — pass through allowed fields only
    const payload = {
      model: model,
      max_tokens: body.max_tokens || 4096,
      messages: body.messages,
    };
    if (body.system) payload.system = body.system;
    if (body.temperature !== undefined) payload.temperature = body.temperature;

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
        headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
        body: result.body,
      };
    } catch (e) {
      context.error('Claude API error:', e.message);
      return { status: 502, headers: CORS_HEADERS, jsonBody: { error: 'Failed to reach Claude API' } };
    }
  },
});
