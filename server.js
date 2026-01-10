/**
 * Gmail Bridge -> PokemonCenter MFA Passcode (SAFE by "after" timestamp)
 *
 * Endpoints:
 *   GET /health
 *   GET /recent?limit=10
 *   GET /code?to=xxx@icloud.com&after=1700000000000
 *
 * Behavior:
 * - /code returns ONLY a passcode from messages that satisfy:
 *     internalDate >= max(after, now - LAST_MINUTES)
 * - This prevents returning an old code when multiple MFA mails exist.
 *
 * SECURITY:
 * - Uses OAuth token.json + credentials json
 * - No passwords stored here
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
require('dotenv').config();

const PORT = Number(process.env.PORT || 8787);

const TOKEN_PATH = path.join(__dirname, process.env.TOKEN_FILE || 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, process.env.CRE_FILE || 'pokemon_cre.json');

const SUBJECT_KEYWORD = 'ログイン用パスコード';

// Strict time window for code (minutes)
const LAST_MINUTES = Number(process.env.LAST_MINUTES || 5);

// Wider search to avoid Gmail indexing delay (minutes)
const QUERY_MINUTES = Number(process.env.QUERY_MINUTES || 60);

// Optional: how many messages to scan for code per request
const CODE_SCAN_LIMIT = Number(process.env.CODE_SCAN_LIMIT || 20);

function sendJson(res, status, obj) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  });
  res.end(JSON.stringify(obj));
}

function getHeader(headers, name) {
  const h = (headers || []).find(x => (x.name || '').toLowerCase() === name.toLowerCase());
  return h ? h.value : '';
}

function extractEmails(headerValue) {
  if (!headerValue) return [];
  return headerValue
    .split(',')
    .map(s => s.trim())
    .map(addr => {
      const match = addr.match(/<([^>]+)>/);
      return (match ? match[1] : addr).trim();
    })
    .filter(Boolean)
    .map(x => x.toLowerCase());
}

function extract6Digits(text) {
  if (!text) return null;
  const norm = String(text).replace(/\r/g, '').replace(/[ \t]+/g, ' ');

  // Best match: 【パスコード】344057
  let m = norm.match(/【\s*パスコード\s*】\s*([0-9]{6})/);
  if (m) return m[1];

  // Variants: パスコード: 344057 / パスコード：344057
  m = norm.match(/パスコード\s*[:：]?\s*([0-9]{6})/);
  if (m) return m[1];

  // Keyword near digits
  m = norm.match(/パスコード[\s\S]{0,240}?([0-9]{6})/);
  if (m) return m[1];

  // Fallback: any 6 digits
  m = norm.match(/\b([0-9]{6})\b/);
  return m ? m[1] : null;
}

function decodeBodyFromPayload(payload) {
  if (!payload) return '';
  const candidates = [];

  // Walk MIME tree and pick text/plain first, then text/html
  function walk(p) {
    if (!p) return;
    const mime = String(p.mimeType || '').toLowerCase();
    if (p.body?.data && (mime === 'text/plain' || mime === 'text/html')) {
      candidates.push({ mime, data: p.body.data });
    }
    if (Array.isArray(p.parts)) p.parts.forEach(walk);
  }

  walk(payload);

  candidates.sort((a, b) => {
    if (a.mime === 'text/plain' && b.mime !== 'text/plain') return -1;
    if (a.mime !== 'text/plain' && b.mime === 'text/plain') return 1;
    return 0;
  });

  if (!candidates.length) return '';

  try {
    // Gmail uses base64url
    const b64 = candidates[0].data.replace(/-/g, '+').replace(/_/g, '/');
    return Buffer.from(b64, 'base64').toString('utf8');
  } catch {
    return '';
  }
}

async function authorize() {
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error(`Missing credentials: ${CREDENTIALS_PATH}`);
  }
  if (!fs.existsSync(TOKEN_PATH)) {
    throw new Error(`Missing token: ${TOKEN_PATH}`);
  }

  const raw = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
  const key = raw.installed || raw.web;
  if (!key) throw new Error('Invalid credentials file (missing installed/web)');

  const { client_id, client_secret, redirect_uris } = key;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);

  const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf-8'));
  oAuth2Client.setCredentials(token);

  return oAuth2Client;
}

async function gmailClient() {
  const auth = await authorize();
  return google.gmail({ version: 'v1', auth });
}

async function listRecent({ limit = 10 }) {
  const gmail = await gmailClient();

  // Broad list (last QUERY_MINUTES minutes), filter by subject in code
  const res = await gmail.users.messages.list({
    userId: 'me',
    q: `newer_than:${QUERY_MINUTES}m`,
    maxResults: Math.min(Math.max(limit, 1), 50),
  });

  const messages = res.data.messages || [];
  const out = [];

  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'metadata',
      metadataHeaders: ['Subject', 'From', 'To', 'Delivered-To', 'X-Original-To', 'Date'],
    });

    const headers = msg.data.payload?.headers || [];
    const subject = (getHeader(headers, 'Subject') || '').trim();

    if (!subject.includes(SUBJECT_KEYWORD)) continue;

    out.push({
      id: m.id,
      internalDate: new Date(Number(msg.data.internalDate || 0)).toISOString(),
      subject,
      from: (getHeader(headers, 'From') || '').trim(),
      to: getHeader(headers, 'To') || '',
      deliveredTo: getHeader(headers, 'Delivered-To') || '',
      xOriginalTo: getHeader(headers, 'X-Original-To') || '',
    });

    if (out.length >= limit) break;
  }

  return out;
}

/**
 * Find MFA code in emails that are newer than BOTH:
 * - afterTs (client started polling)
 * - now - LAST_MINUTES
 */
async function getCode({ toEmail = '', afterTs = 0 }) {
  const gmail = await gmailClient();

  const now = Date.now();
  const want = String(toEmail || '').toLowerCase().trim();

  // Lower bound by "after" (from client) and LAST_MINUTES
  const lastWindowTs = now - LAST_MINUTES * 60 * 1000;
  const minTs = Math.max(Number(afterTs || 0), lastWindowTs);

  // Search by subject, widen time window to handle indexing delays
  const q = `newer_than:${QUERY_MINUTES}m subject:("${SUBJECT_KEYWORD}")`;

  const listRes = await gmail.users.messages.list({
    userId: 'me',
    q,
    maxResults: CODE_SCAN_LIMIT,
  });

  const messages = listRes.data.messages || [];
  if (!messages.length) return { found: false, code: null, reason: 'no_messages' };

  // Gmail list is newest-first typically, so first match is best
  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'full',
    });

    const internalDate = Number(msg.data.internalDate || 0);

    // Skip old messages (before minTs)
    if (internalDate && internalDate < minTs) continue;

    const headers = msg.data.payload?.headers || [];
    const subject = (getHeader(headers, 'Subject') || '').trim();
    if (!subject.includes(SUBJECT_KEYWORD)) continue;

    // Optional recipient filter: check multiple headers
    if (want) {
      const toHeader = getHeader(headers, 'To');
      const deliveredTo = getHeader(headers, 'Delivered-To');
      const xOriginalTo = getHeader(headers, 'X-Original-To');

      const addrPool = [
        ...extractEmails(toHeader),
        ...extractEmails(deliveredTo),
        ...extractEmails(xOriginalTo),
      ];

      // If there is recipient info and it doesn't match, skip
      if (addrPool.length && !addrPool.includes(want)) continue;
    }

    // Try snippet first
    const snippet = msg.data.snippet || '';
    let code = extract6Digits(snippet);
    if (code) {
      return { found: true, code, where: 'snippet', internalDate };
    }

    // Then decode body
    const bodyText = decodeBodyFromPayload(msg.data.payload);
    code = extract6Digits(bodyText);
    if (code) {
      return { found: true, code, where: 'body', internalDate };
    }
  }

  return {
    found: false,
    code: null,
    reason: `no_code_after_${new Date(minTs).toISOString()}`,
    minTs,
  };
}

http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') return sendJson(res, 204, { ok: true });

  const u = new URL(req.url, `http://${req.headers.host}`);

  try {
    if (u.pathname === '/health') {
      return sendJson(res, 200, { ok: true });
    }

    if (u.pathname === '/recent') {
      const limit = Number(u.searchParams.get('limit') || '10');
      const list = await listRecent({ limit });
      return sendJson(res, 200, { ok: true, count: list.length, list });
    }

    if (u.pathname === '/code') {
      const to = u.searchParams.get('to') || '';
      const after = Number(u.searchParams.get('after') || '0'); // ms timestamp
      const r = await getCode({ toEmail: to, afterTs: after });
      return sendJson(res, 200, { ok: true, ...r });
    }

    return sendJson(res, 404, { ok: false, error: 'Not found' });
  } catch (e) {
    return sendJson(res, 500, { ok: false, error: String(e?.message || e) });
  }
}).listen(PORT, '127.0.0.1', () => {
  console.log(`Gmail bridge running: http://127.0.0.1:${PORT}`);
  console.log(`Test: /health , /recent?limit=10 , /code?to=xxx@icloud.com&after=${Date.now()}`);
  console.log(`Window: code=${LAST_MINUTES}m, query=${QUERY_MINUTES}m`);
});
