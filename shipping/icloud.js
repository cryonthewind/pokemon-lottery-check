// NOTE: Code comments are in English as requested.

// iCloud IMAP: export Pokemon Center "Shipping" to EXCEL (NON-UNIQUE)
// Subject: 【ポケモンセンターオンライン】商品が出荷されました
//
// Excel columns (ONLY):
// Date, To, WaybillNo, Product Name, Price, Address, TrackingUrl
//
// Rules:
// - No unique aggregation
// - If From is hotmail/outlook => use From email as "To"
// - Otherwise => use To email (first address)
// - "To" is pure email only
// - Parse 【商品情報】 line like:
//   9900000007003 【抽選販売】XXXX 5,400円 1個
//   => Product Name: XXXX, Price: 5,400
//
// SECURITY:
// - Do NOT hardcode iCloud credentials. Use environment variables (.env).

const { ImapFlow } = require('imapflow');
const ExcelJS = require('exceljs');
const path = require('path');
require('dotenv').config();

// -------------------- Config --------------------
const SUBJECT_KEY = '【ポケモンセンターオンライン】商品が出荷されました';
const DAYS_BACK = Number(process.env.DAYS_BACK || 7);
// ------------------------------------------------

function pct(n, d) {
  if (!d) return '0.00%';
  return `${((n / d) * 100).toFixed(2)}%`;
}

// Decide which mailbox to store in "To" column:
// - If sender is hotmail/outlook => use FROM email
// - Otherwise => use TO email (first)
function decideToEmailForRow(fromAddr, toAddr) {
  const fromEmail = (fromAddr && fromAddr.address) ? String(fromAddr.address).trim() : '';
  const toEmail = (toAddr && toAddr.address) ? String(toAddr.address).trim() : '';
  const isHotmail = /hotmail\.com|outlook\.com/i.test(fromEmail);
  return isHotmail ? fromEmail : toEmail;
}

// ------------------------------------------------------
// MIME helpers (minimal but practical)
// ------------------------------------------------------

// Decode quoted-printable (best-effort) + remove soft line breaks
function decodeQuotedPrintable(input) {
  if (!input) return '';
  // Remove soft line breaks "=\r\n" or "=\n"
  let s = input.replace(/=\r?\n/g, '');
  // Decode =XX hex
  s = s.replace(/=([0-9A-Fa-f]{2})/g, (_, hex) => {
    return String.fromCharCode(parseInt(hex, 16));
  });
  return s;
}

// Decode base64 (best-effort)
function decodeBase64(input) {
  if (!input) return '';
  // Remove whitespace/newlines
  const s = input.replace(/\s+/g, '');
  try {
    return Buffer.from(s, 'base64').toString('utf-8');
  } catch {
    // Fallback: return as-is if decode fails
    return input;
  }
}

// Detect & decode transfer encoding for a part body
function decodeByTransferEncoding(body, transferEncoding) {
  const enc = String(transferEncoding || '').toLowerCase().trim();
  if (enc === 'base64') return decodeBase64(body);
  if (enc === 'quoted-printable') return decodeQuotedPrintable(body);
  return body || '';
}

// Extract boundary from Content-Type header line
function extractBoundary(contentTypeLine) {
  if (!contentTypeLine) return '';
  const m = contentTypeLine.match(/boundary="?([^";]+)"?/i);
  return m ? m[1] : '';
}

// Parse headers block into map (lowercased keys)
function parseHeaderBlock(headerText) {
  const lines = headerText.split(/\r?\n/);
  // Handle folded headers
  const unfolded = [];
  for (const line of lines) {
    if (/^\s/.test(line) && unfolded.length) {
      unfolded[unfolded.length - 1] += ' ' + line.trim();
    } else {
      unfolded.push(line.trim());
    }
  }

  const map = {};
  for (const l of unfolded) {
    const idx = l.indexOf(':');
    if (idx === -1) continue;
    const k = l.slice(0, idx).trim().toLowerCase();
    const v = l.slice(idx + 1).trim();
    map[k] = v;
  }
  return map;
}

// Minimal MIME part extractor: returns { plainText, htmlText }
function extractTextFromRfc822(rawRfc822) {
  if (!rawRfc822) return { plainText: '', htmlText: '' };

  const raw = rawRfc822.replace(/\r\n/g, '\n');

  // Split top headers / body
  const splitIdx = raw.indexOf('\n\n');
  const topHeadersText = splitIdx !== -1 ? raw.slice(0, splitIdx) : '';
  const bodyAll = splitIdx !== -1 ? raw.slice(splitIdx + 2) : raw;

  const topHeaders = parseHeaderBlock(topHeadersText);
  const topContentType = topHeaders['content-type'] || '';
  const boundary = extractBoundary(topContentType);

  // If not multipart, just decode body best-effort
  if (!/multipart\//i.test(topContentType) || !boundary) {
    const transfer = topHeaders['content-transfer-encoding'] || '';
    const decoded = decodeByTransferEncoding(bodyAll, transfer);
    return { plainText: decoded, htmlText: '' };
  }

  // Multipart: split by boundary markers
  const boundaryMark = '--' + boundary;
  const endMark = boundaryMark + '--';

  const parts = bodyAll.split(boundaryMark)
    .map(p => p.replace(/^\n/, ''))
    .filter(p => p && p.trim() && !p.startsWith('--') && !p.startsWith(endMark));

  let plainText = '';
  let htmlText = '';

  for (const part of parts) {
    const i = part.indexOf('\n\n');
    const headerText = i !== -1 ? part.slice(0, i) : '';
    const bodyText = i !== -1 ? part.slice(i + 2) : part;

    const h = parseHeaderBlock(headerText);
    const ct = (h['content-type'] || '').toLowerCase();
    const te = h['content-transfer-encoding'] || '';

    const decoded = decodeByTransferEncoding(bodyText, te);

    // Some servers embed multipart/alternative inside multipart/mixed.
    // We handle nested multipart shallowly by recursion if needed.
    if (ct.includes('multipart/') && extractBoundary(h['content-type'] || '')) {
      const nested = extractTextFromRfc822(headerText + '\n\n' + bodyText);
      if (!plainText && nested.plainText) plainText = nested.plainText;
      if (!htmlText && nested.htmlText) htmlText = nested.htmlText;
      continue;
    }

    if (!plainText && ct.includes('text/plain')) plainText = decoded;
    if (!htmlText && ct.includes('text/html')) htmlText = decoded;
  }

  return { plainText, htmlText };
}

// Light HTML -> text
function htmlToText(html) {
  if (!html) return '';
  return html
    .replace(/\r\n/g, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"');
}

// ------------------------------------------------------
// Business parsers
// ------------------------------------------------------

function cleanProductName(name) {
  if (!name) return '';
  let s = name.trim();
  s = s.replace(/^【抽選販売】\s*/g, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// Extract product lines in 【商品情報】 for shipping format (no 小計)
function extractProductLinesShipping(fullText) {
  if (!fullText) return [];
  const t = fullText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  const startIdx = t.indexOf('【商品情報】');
  if (startIdx === -1) return [];

  const endMarkers = [
    '【お届け先情報】',
    '【注文者情報】',
    '【ご請求金額】',
    '【お支払い情報】',
    '【配送情報】',
    '【注文情報】',
    '【注意事項】',
    '【お届け先】',
    'お届け先',
    '配送情報',
  ];

  let endIdx = t.length;
  for (const m of endMarkers) {
    const i = t.indexOf(m, startIdx + 1);
    if (i !== -1 && i < endIdx) endIdx = i;
  }

  const block = t.slice(startIdx, endIdx);

  const lines = block
    .split('\n')
    .map(l => (l || '').trim())
    .filter(Boolean);

  // Example line contains: JAN + name + price円 + qty個
  return lines.filter(l => /\b\d{8,14}\b/.test(l) && /円/.test(l) && /個/.test(l));
}

function parseShippingProductLine(line) {
  const janMatch = line.match(/\b(\d{8,14})\b/);
  const qtyMatch = line.match(/(\d+)\s*個/);
  const priceMatch = line.match(/([0-9,]+)\s*円/);

  const jan = janMatch ? janMatch[1] : '';
  const qty = qtyMatch ? qtyMatch[1] : '';
  const price = priceMatch ? priceMatch[1] : '';

  let name = line;
  if (jan) name = name.replace(jan, '').trim();
  if (price) name = name.replace(new RegExp(`${price}\\s*円`), '').trim();
  if (qty) name = name.replace(new RegExp(`${qty}\\s*個`), '').trim();

  name = name.replace(/^[-:：\s]+/, '').trim();
  name = cleanProductName(name);

  return { jan, name, price, raw: line };
}

function extractProductAndPriceForShipping(fullText) {
  const lines = extractProductLinesShipping(fullText);
  if (!lines.length) return { productName: '', price: '' };

  const names = [];
  const prices = [];

  for (const line of lines) {
    const p = parseShippingProductLine(line);
    if (p.name) names.push(p.name);
    if (p.price) prices.push(p.price);
  }

  // Deduplicate while preserving order
  function uniq(arr) {
    const seen = new Set();
    const out = [];
    for (const x of arr) {
      if (!x) continue;
      if (seen.has(x)) continue;
      seen.add(x);
      out.push(x);
    }
    return out;
  }

  return {
    productName: uniq(names).join(' / '),
    price: uniq(prices).join(' / '),
  };
}

// Extract tracking URL (more flexible + join broken lines)
function extractTrackingUrl(fullText) {
  if (!fullText) return '';

  // Remove quoted-printable soft breaks in case they remain
  const t = fullText.replace(/=\r?\n/g, '');

  // Grab first URL starting with https://member.kms.
  const m = t.match(/https?:\/\/member\.kms\.[^\s<>"']+/i);
  if (!m) return '';

  // Clean trailing punctuation
  return m[0].replace(/[),.]+$/, '');
}

// Extract waybill number by labels (best-effort)
function extractWaybillNo(fullText, trackingUrl) {
  if (!fullText) fullText = '';
  const t = fullText.replace(/=\r?\n/g, '');

  let m = t.match(/(送り状番号|お問い合わせ伝票番号)[：:]\s*([0-9\-]{5,})/);
  if (m) return m[2];

  m = t.match(/(送り状番号|お問い合わせ伝票番号)[^0-9]{0,50}([0-9\-]{5,})/);
  if (m) return m[2];

  m = t.match(
    /(送り状番号|お問い合わせ伝票番号)[^0-9\r\n]{0,10}[\r\n]+[^\r\n]*?([0-9\-]{5,})/
  );
  if (m) return m[2];

  // Fallback: if URL has pno=...
  if (trackingUrl) {
    const u = trackingUrl.replace(/=\r?\n/g, '');
    const um = u.match(/pno=([0-9\-]{5,})/i);
    if (um) return um[1];
  }

  return '';
}

// Extract shipping address block after お届け先
function extractAddress(fullText) {
  if (!fullText) return '';

  const t = fullText.replace(/=\r?\n/g, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = t.split('\n').map(l => (l || '').trim());

  let idx = lines.findIndex(line => line.includes('お届け先'));
  if (idx === -1) return '';

  let name = '';
  let zip = '';
  const addressLines = [];

  for (let i = idx + 1; i < lines.length; i++) {
    const line = (lines[i] || '').trim();

    if (
      line.startsWith('【') ||
      line.includes('お支払い方法') ||
      line.includes('ご注文商品') ||
      line.includes('ご注文内容') ||
      line.includes('配送情報') ||
      line.includes('注文情報')
    ) {
      break;
    }

    if (!line) continue;

    if (!name && line.includes('様')) {
      name = line;
      continue;
    }

    if (!zip && line.startsWith('〒')) {
      zip = line;
      continue;
    }

    addressLines.push(line);
  }

  const parts = [];
  if (name) parts.push(name);
  if (zip) parts.push(zip);
  if (addressLines.length > 0) parts.push(addressLines.join(' '));

  return parts.join(' ');
}

// ------------------------------------------------------
// Export to Excel
// ------------------------------------------------------
async function exportToExcel(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'icloud-pokemon-shipping-export';
  workbook.created = new Date();

  const ws = workbook.addWorksheet('Shipping');
  ws.columns = [
    { header: 'Date', key: 'date', width: 28 },
    { header: 'To', key: 'to', width: 30 },
    { header: 'WaybillNo', key: 'waybillNo', width: 16 },
    { header: 'Product Name', key: 'productName', width: 70 },
    { header: 'Price', key: 'price', width: 12 },
    { header: 'Address', key: 'address', width: 55 },
    { header: 'TrackingUrl', key: 'trackingUrl', width: 55 },
  ];

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.getRow(1).font = { bold: true };
  ws.autoFilter = { from: 'A1', to: 'G1' };

  rows.forEach(r => ws.addRow(r));

  // Make URL clickable
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const cell = row.getCell('trackingUrl');
    const url = cell.value;
    if (typeof url === 'string' && url.startsWith('http')) {
      cell.value = { text: url, hyperlink: url };
    }
  });

  // Wrap long text
  ws.getColumn('productName').alignment = { wrapText: true, vertical: 'top' };
  ws.getColumn('address').alignment = { wrapText: true, vertical: 'top' };
  ws.getColumn('trackingUrl').alignment = { wrapText: true, vertical: 'top' };

  const outPath = path.join(__dirname, 'icloud_pokemon_shipping.xlsx');
  await workbook.xlsx.writeFile(outPath);

  console.log(`\nExcel exported: ${outPath}`);
}

// ------------------------------------------------------
// Main
// ------------------------------------------------------
async function listIcloudPokemonShipping() {
  const ICLOUD_USER = process.env.ICLOUD_USER;
  const ICLOUD_PASS = process.env.ICLOUD_APP_PASSWORD;

  if (!ICLOUD_USER || !ICLOUD_PASS) {
    console.error('Missing ICLOUD_USER / ICLOUD_APP_PASSWORD in .env');
    process.exit(1);
  }

  const client = new ImapFlow({
    host: 'imap.mail.me.com',
    port: 993,
    secure: true,
    auth: {
      user: ICLOUD_USER,
      pass: ICLOUD_PASS,
    },
    disableCompression: true,
    tls: {
      // WARNING: for testing only. Remove this if TLS works without it.
      rejectUnauthorized: false,
    },
    logger: false,
  });

  try {
    await client.connect();
    await client.mailboxOpen('INBOX');

    const sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - DAYS_BACK);

    const messageUids = await client.search({ since: sinceDate });

    console.log('=====================');
    console.log(`Search since: ${sinceDate.toISOString()}`);
    console.log(`Matched UIDs: ${messageUids.length}`);
    console.log('=====================');

    const rows = [];
    const mailLogs = [];

    let matchedMessages = 0;
    let noProductFound = 0;
    let noAddressFound = 0;
    let noWaybillFound = 0;
    let noUrlFound = 0;

    for await (const msg of client.fetch(messageUids, { envelope: true, source: true })) {
      const subject = msg.envelope?.subject || '';
      if (!subject.includes(SUBJECT_KEY)) continue;
      matchedMessages++;

      const fromAddr = (msg.envelope?.from || [])[0] || null;
      const toAddr = (msg.envelope?.to || [])[0] || null;

      const fromEmail = fromAddr?.address ? String(fromAddr.address).trim() : '';
      const toEmailRaw = toAddr?.address ? String(toAddr.address).trim() : '';
      const toEmail = decideToEmailForRow(fromAddr, toAddr) || toEmailRaw || '';

      const dateObj = msg.envelope?.date ? new Date(msg.envelope.date) : null;
      const dateStr = dateObj && !Number.isNaN(dateObj.getTime())
        ? dateObj.toUTCString()
        : '';

      const raw = msg.source
        ? (Buffer.isBuffer(msg.source) ? msg.source.toString('utf-8') : String(msg.source))
        : '';

      // ✅ Properly extract decoded text/plain or html->text
      const { plainText, htmlText } = extractTextFromRfc822(raw);
      const bodyText = (plainText && plainText.trim())
        ? plainText
        : htmlToText(htmlText || '');

      const trackingUrl = extractTrackingUrl(bodyText);
      const waybillNo = extractWaybillNo(bodyText, trackingUrl);
      const address = extractAddress(bodyText);
      const { productName, price } = extractProductAndPriceForShipping(bodyText);

      if (!trackingUrl) noUrlFound++;
      if (!waybillNo) noWaybillFound++;
      if (!address) noAddressFound++;
      if (!productName) noProductFound++;

      mailLogs.push({
        date: dateStr,
        to: toEmail,
        from: fromEmail,
        hasProduct: productName ? 'YES' : 'NO',
        hasAddr: address ? 'YES' : 'NO',
        hasWaybill: waybillNo ? 'YES' : 'NO',
        hasUrl: trackingUrl ? 'YES' : 'NO',
      });

      rows.push({
        date: dateStr,
        to: toEmail,
        waybillNo: waybillNo || '',
        productName: productName || '',
        price: price || '',
        address: address || '',
        trackingUrl: trackingUrl || '',
      });
    }

    console.log('\n========== MAIL LIST ==========');
    console.log(`Matched messages: ${matchedMessages}`);
    for (const it of mailLogs) {
      console.log(
        `date=${it.date || 'N/A'} | to=${it.to || 'N/A'} | from=${it.from || 'N/A'} | product=${it.hasProduct} | addr=${it.hasAddr} | waybill=${it.hasWaybill} | url=${it.hasUrl}`
      );
    }
    console.log('===============================');

    console.log('\n========== EXPORT STATS ==========');
    console.log(`Excel rows: ${rows.length}`);
    console.log(`NO product: ${noProductFound}`);
    console.log(`NO address: ${noAddressFound}`);
    console.log(`NO waybill: ${noWaybillFound}`);
    console.log(`NO url: ${noUrlFound}`);
    console.log('=================================');

    await exportToExcel(rows);

    console.log('\n========== DONE ==========');
    console.log(`Hit rate (matched/UIDs): ${pct(matchedMessages, messageUids.length)} (${matchedMessages}/${messageUids.length})`);
    console.log('==========================');

  } catch (err) {
    console.error(err);
  } finally {
    if (!client.closed) {
      await client.logout().catch(() => {});
    }
  }
}

listIcloudPokemonShipping();
