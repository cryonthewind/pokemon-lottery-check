/**
 * iCloud IMAP Bridge -> PokemonCenter MFA Passcode (SAFE by "after" timestamp)
 *
 * Endpoints:
 *   GET /health
 *   GET /recent?limit=10
 *   GET /code?to=xxx@icloud.com&after=1700000000000
 *
 * Behavior:
 * - /code returns ONLY a passcode from messages that satisfy:
 *     internalDate >= max(after, now - LAST_MINUTES)
 *
 * Notes:
 * - iCloud has no Gmail-like REST API, so we use IMAP.
 * - Uses App-Specific Password (recommended when 2FA enabled).
 */

const http = require('http');
require('dotenv').config();

const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

const PORT = Number(process.env.PORT || 8787);

// iCloud IMAP config
const ICLOUD_USER = process.env.ICLOUD_USER || ''; // e.g. xxx@icloud.com
const ICLOUD_APP_PASSWORD = process.env.ICLOUD_APP_PASSWORD || ''; // app-specific password
const ICLOUD_HOST = process.env.ICLOUD_HOST || 'imap.mail.me.com';
const ICLOUD_PORT = Number(process.env.ICLOUD_PORT || 993);
const ICLOUD_SECURE = String(process.env.ICLOUD_SECURE || 'true') === 'true';
const ICLOUD_MAILBOX = process.env.ICLOUD_MAILBOX || 'INBOX';

// Subject keyword used by PokemonCenter
const SUBJECT_KEYWORD = process.env.SUBJECT_KEYWORD || 'ログイン用パスコード';

// Strict time window for code (minutes)
const LAST_MINUTES = Number(process.env.LAST_MINUTES || 5);

// Wider scan window (minutes) to avoid missing due to delays
const QUERY_MINUTES = Number(process.env.QUERY_MINUTES || 60);

// How many newest messages to scan for code
const CODE_SCAN_LIMIT = Number(process.env.CODE_SCAN_LIMIT || 20);

// Debug logs
const DEBUG = String(process.env.DEBUG || '1') === '1';

// TLS verify switch (ONLY for debugging when you have cert issues)
const IMAP_REJECT_UNAUTHORIZED =
  String(process.env.IMAP_REJECT_UNAUTHORIZED || 'true') === 'true';

function log(...args) {
  if (DEBUG) console.log(...args);
}

function sendJson(res, status, obj) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  });
  res.end(JSON.stringify(obj));
}

function extractEmails(headerValue) {
  if (!headerValue) return [];
  return String(headerValue)
    .split(',')
    .map(s => s.trim())
    .map(addr => {
      const match = addr.match(/<([^>]+)>/);
      return (match ? match[1] : addr).trim();
    })
    .filter(Boolean)
    .map(x => x.toLowerCase());
}

function normalizeToList(addressObjOrText) {
  // mailparser gives address objects; sometimes we only have raw strings
  if (!addressObjOrText) return [];
  if (typeof addressObjOrText === 'string') return extractEmails(addressObjOrText);

  // AddressObject: { value: [{ address, name }], text }
  if (addressObjOrText.value && Array.isArray(addressObjOrText.value)) {
    return addressObjOrText.value
      .map(v => (v.address || '').toLowerCase().trim())
      .filter(Boolean);
  }

  if (addressObjOrText.text) return extractEmails(addressObjOrText.text);
  return [];
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

function toIso(ms) {
  try {
    return new Date(Number(ms || 0)).toISOString();
  } catch {
    return null;
  }
}

function fmtAddrList(list) {
  if (!Array.isArray(list) || !list.length) return '';
  const a = list[0];
  if (!a) return '';
  const name = a.name || '';
  const addr = a.address || '';
  if (name && addr) return `${name} <${addr}>`;
  return addr || name || '';
}

async function withImap(fn) {
  if (!ICLOUD_USER || !ICLOUD_APP_PASSWORD) {
    throw new Error('Missing ICLOUD_USER or ICLOUD_APP_PASSWORD in .env');
  }

  const client = new ImapFlow({
    host: ICLOUD_HOST,
    port: ICLOUD_PORT,
    secure: ICLOUD_SECURE,
    auth: { user: ICLOUD_USER, pass: ICLOUD_APP_PASSWORD },
    // TLS options
    tls: {
      // WARNING: keep true for security; set false only if you must debug cert issues
      rejectUnauthorized: IMAP_REJECT_UNAUTHORIZED,
      servername: ICLOUD_HOST,
    },
    logger: false,
  });

  await client.connect();
  try {
    await client.mailboxOpen(ICLOUD_MAILBOX);
    return await fn(client);
  } finally {
    try {
      await client.logout();
    } catch {
      // ignore
    }
  }
}

async function listRecent({ limit = 10 }) {
  const now = Date.now();
  const since = new Date(now - QUERY_MINUTES * 60 * 1000);

  return await withImap(async client => {
    // IMAP SINCE is day-level in search, then we filter precisely by internalDate
    const uids = await client.search({ since });

    // Newest-first scan; take more than limit because we filter by subject
    const scan = uids.slice(-Math.min(Math.max(limit * 5, 20), 300)).reverse();

    const out = [];
    for (const uid of scan) {
      const msg = await client.fetchOne(uid, { envelope: true, internalDate: true });
      const subject = (msg.envelope?.subject || '').trim();
      if (!subject.includes(SUBJECT_KEYWORD)) continue;

      const internalDateMs = msg.internalDate ? new Date(msg.internalDate).getTime() : 0;

      out.push({
        uid,
        internalDate: toIso(internalDateMs),
        subject,
        from: fmtAddrList(msg.envelope?.from),
        to: fmtAddrList(msg.envelope?.to),
      });

      if (out.length >= limit) break;
    }

    return out;
  });
}

/**
 * Find MFA code in emails that are newer than BOTH:
 * - afterTs (client started polling)
 * - now - LAST_MINUTES
 */
async function getCode({ toEmail = '', afterTs = 0 }) {
  const now = Date.now();
  const want = String(toEmail || '').toLowerCase().trim();

  let after = Number(afterTs || 0);
  if (after > now + 60_000) {
    log('[WARN] afterTs is in the future, clamping', { after, now });
    after = now;
  }

  const lastWindowTs = now - LAST_MINUTES * 60 * 1000;
  const minTs = Math.max(after, lastWindowTs);

  const since = new Date(now - QUERY_MINUTES * 60 * 1000);

  return await withImap(async client => {
    const uids = await client.search({ since });
    if (!uids.length) return { found: false, code: null, reason: 'no_messages' };

    const newest = uids.slice(-Math.min(CODE_SCAN_LIMIT, uids.length)).reverse();

    log('[INFO] scan_start', {
      want: want || '(none)',
      uidsTotal: uids.length,
      scanCount: newest.length,
      minTs,
      minTsIso: toIso(minTs),
      after,
      afterIso: toIso(after),
      lastWindowTs,
      lastWindowIso: toIso(lastWindowTs),
    });

    for (const uid of newest) {
      const meta = await client.fetchOne(uid, { envelope: true, internalDate: true });

      const internalDateMs = meta.internalDate ? new Date(meta.internalDate).getTime() : 0;
      const subject = (meta.envelope?.subject || '').trim();

      log('[SCAN]', { uid, internalDateMs, internalDateIso: toIso(internalDateMs), subject });

      if (internalDateMs && internalDateMs < minTs) {
        log('[SKIP] too_old', { uid, internalDateMs, minTs });
        continue;
      }

      if (!subject.includes(SUBJECT_KEYWORD)) {
        log('[SKIP] subject_mismatch', { uid, subject });
        continue;
      }

      // ---- Recipient gating ----
      // If "to" is provided, we enforce recipient match.
      // iCloud IMAP ENVELOPE often has empty "to", so we fallback to parsed headers.
      if (want) {
        const toList = normalizeToList(meta.envelope?.to);
        const ccList = normalizeToList(meta.envelope?.cc);
        const bccList = normalizeToList(meta.envelope?.bcc);
        const pool = [...toList, ...ccList, ...bccList].filter(Boolean);

        log('[INFO] rcpt_pool_envelope', { uid, want, pool });

        if (pool.length) {
          if (!pool.includes(want)) {
            log('[SKIP] recipient_mismatch_envelope', { uid, want, pool });
            continue;
          }
        } else {
          // Fallback: download and parse headers to check recipient
          const dl0 = await client.download(uid);
          const parsed0 = await simpleParser(dl0.content);

          const parsedTo = normalizeToList(parsed0.to);
          const parsedCc = normalizeToList(parsed0.cc);
          const parsedBcc = normalizeToList(parsed0.bcc);

          // Also check raw headers that may appear in forwarded/relay emails
          const hdr = parsed0.headers || new Map();
          const deliveredTo = String(hdr.get('delivered-to') || '').toLowerCase();
          const xOriginalTo = String(hdr.get('x-original-to') || '').toLowerCase();
          const toHeader = String(hdr.get('to') || '').toLowerCase();

          const pool2 = [
            ...parsedTo,
            ...parsedCc,
            ...parsedBcc,
            ...extractEmails(toHeader),
            ...extractEmails(deliveredTo),
            ...extractEmails(xOriginalTo),
          ].filter(Boolean);

          log('[INFO] rcpt_pool_parsed', {
            uid,
            want,
            pool2,
            deliveredTo,
            xOriginalTo,
            toHeader,
          });

          // If we still cannot determine recipient, reject to avoid returning wrong code
          if (!pool2.length) {
            log('[SKIP] recipient_unknown_reject', { uid, want });
            continue;
          }

          if (!pool2.includes(want)) {
            log('[SKIP] recipient_mismatch_parsed', { uid, want, pool2 });
            continue;
          }

          // If recipient matched in parsed headers, reuse parsed0 for code extraction
          const text0 = parsed0.text || '';
          const html0 = parsed0.html ? String(parsed0.html) : '';
          const combined0 = `${text0}\n${html0}`;

          let code0 = extract6Digits(text0);
          if (code0) {
            log('[HIT] code_from_text', { uid, code: code0 });
            return { found: true, code: code0, where: 'text', internalDate: internalDateMs };
          }

          code0 = extract6Digits(combined0);
          if (code0) {
            log('[HIT] code_from_parsed', { uid, code: code0 });
            return { found: true, code: code0, where: 'parsed', internalDate: internalDateMs };
          }

          log('[MISS] no_code_in_message', { uid });
          continue;
        }
      }

      // If no "to" filter, or recipient matched in envelope, then parse mail for OTP
      const dl = await client.download(uid);
      const parsed = await simpleParser(dl.content);

      const text = parsed.text || '';
      const html = parsed.html ? String(parsed.html) : '';
      const combined = `${text}\n${html}`;

      let code = extract6Digits(text);
      if (code) {
        log('[HIT] code_from_text', { uid, code });
        return { found: true, code, where: 'text', internalDate: internalDateMs };
      }

      code = extract6Digits(combined);
      if (code) {
        log('[HIT] code_from_parsed', { uid, code });
        return { found: true, code, where: 'parsed', internalDate: internalDateMs };
      }

      log('[MISS] no_code_in_message', { uid });
    }

    return {
      found: false,
      code: null,
      reason: `no_code_after_${toIso(minTs)}`,
      minTs,
      minTsIso: toIso(minTs),
    };
  });
}


http
  .createServer(async (req, res) => {
    if (req.method === 'OPTIONS') return sendJson(res, 204, { ok: true });

    const u = new URL(req.url, `http://${req.headers.host}`);

    try {
      if (u.pathname === '/health') {
        return sendJson(res, 200, {
          ok: true,
          provider: 'icloud-imap',
          mailbox: ICLOUD_MAILBOX,
          subjectKeyword: SUBJECT_KEYWORD,
          windowMinutes: { code: LAST_MINUTES, query: QUERY_MINUTES },
          tls: { rejectUnauthorized: IMAP_REJECT_UNAUTHORIZED },
        });
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
      log('[ERROR]', e);
      return sendJson(res, 500, { ok: false, error: String(e?.message || e) });
    }
  })
  .listen(PORT, '127.0.0.1', () => {
    console.log(`iCloud bridge running: http://127.0.0.1:${PORT}`);
    console.log(`Test: /health , /recent?limit=10 , /code?to=xxx@icloud.com&after=${Date.now()}`);
    console.log(`Window: code=${LAST_MINUTES}m, query=${QUERY_MINUTES}m`);
    console.log(`TLS rejectUnauthorized: ${IMAP_REJECT_UNAUTHORIZED}`);
  });
