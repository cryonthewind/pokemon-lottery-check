// Gmail: export Pokemon Center "Order Completed" emails to EXCEL (NON-UNIQUE)
// Subject: [ポケモンセンターオンライン]注文完了のお知らせ
//
// Excel columns (ONLY):
// Date, To, JAN, Product Name, Qty, Subtotal
//
// Rules:
// - No unique aggregation
// - If From is hotmail/outlook => use From email as "To" (mapping key)
// - Otherwise => use To email as "To"
// - "To" column must be pure email only (no display name)
// - Product Name: remove leading 【抽選販売】 and remove trailing schedule brackets like 【2026...発送予定】

const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'pokemon_cre.json');

function getHeader(headers, name) {
  const h = headers.find(h => (h.name || '').toLowerCase() === name.toLowerCase());
  return h ? (h.value || '') : '';
}

// Extract emails from header like:
// 'Name <a@b.com>, "X" <c@d.com>' => ['a@b.com','c@d.com']
// 'a@b.com' => ['a@b.com']
function extractEmails(headerValue) {
  if (!headerValue) return [];
  return headerValue
    .split(',')
    .map(s => s.trim())
    .map(addr => {
      const match = addr.match(/<([^>]+)>/);
      return (match ? match[1] : addr).trim();
    })
    .filter(Boolean);
}

// Decide which header to use for the "To" column:
// - If sender is hotmail/outlook => use FROM email
// - Otherwise => use TO email
function decideToEmailForRow(fromHeader, toHeader) {
  const fromEmails = extractEmails(fromHeader);
  const toEmails = extractEmails(toHeader);

  const isHotmail = /hotmail\.com|outlook\.com/i.test(fromHeader || '');

  if (isHotmail) return fromEmails[0] || ''; // only email
  return toEmails[0] || ''; // only email
}

// ===================== Gmail body decode helpers =====================
// Decode base64url (Gmail uses base64url)
function decodeBase64Url(data) {
  if (!data) return '';
  const b64 = data.replace(/-/g, '+').replace(/_/g, '/');
  const pad = b64.length % 4 ? '='.repeat(4 - (b64.length % 4)) : '';
  return Buffer.from(b64 + pad, 'base64').toString('utf-8');
}

// Walk through payload parts and collect text/plain (preferred) and text/html (fallback)
function collectBodyTexts(payload) {
  const out = { plain: [], html: [] };

  function walk(part) {
    if (!part) return;

    const mime = (part.mimeType || '').toLowerCase();
    const bodyData = part.body && part.body.data ? part.body.data : '';

    if (mime === 'text/plain' && bodyData) out.plain.push(decodeBase64Url(bodyData));
    if (mime === 'text/html' && bodyData) out.html.push(decodeBase64Url(bodyData));

    const parts = part.parts || [];
    for (const p of parts) walk(p);
  }

  walk(payload);
  return out;
}

// Very light HTML to text fallback (not perfect, but practical)
function htmlToText(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"');
}

// Extract product lines under 【商品情報】 and keep lines containing "小計"
function extractProductLines(fullText) {
  if (!fullText) return [];

  const startIdx = fullText.indexOf('【商品情報】');
  if (startIdx === -1) return [];

  const endMarkers = [
    '【お届け先情報】',
    '【注文者情報】',
    '【ご請求金額】',
    '【お支払い情報】',
    '【配送情報】',
    '【注文情報】',
    '【注意事項】',
  ];

  let endIdx = fullText.length;
  for (const m of endMarkers) {
    const i = fullText.indexOf(m, startIdx + 1);
    if (i !== -1 && i < endIdx) endIdx = i;
  }

  const block = fullText.slice(startIdx, endIdx);

  const lines = block
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  // Product lines usually contain 小計
  return lines.filter(l => l.includes('小計'));
}

// Clean product name:
// - remove leading 【抽選販売】
// - remove trailing schedule bracket like 【2026年...発送予定】
// - also trim spaces
function cleanProductName(name) {
  if (!name) return '';

  let s = name.trim();

  // remove leading bracket tag
  s = s.replace(/^【抽選販売】\s*/g, '');

  // remove trailing 【...発送予定】 or 【...発送】 etc (best-effort)
  // Example: 【2026年2月中旬発送予定】, 【2026年2月15日（日）～28日（土）発送予定】
  s = s.replace(/【[^【】]*発送[^【】]*】\s*$/g, '');

  // normalize spaces
  s = s.replace(/\s+/g, ' ').trim();

  return s;
}

// Parse one product line into fields (best-effort)
function parseProductLine(line) {
  // Example raw:
  // 9900000006808 【抽選販売】ポケモンカードゲーム ... BOX【2026...発送予定】 (1個) 小計 5,500円
  const janMatch = line.match(/\b(\d{8,14})\b/);
  const qtyMatch = line.match(/\((\d+)\s*個\)/);
  const subtotalMatch = line.match(/小計\s*([0-9,]+円)/);

  const jan = janMatch ? janMatch[1] : '';
  const qty = qtyMatch ? qtyMatch[1] : '';
  const subtotal = subtotalMatch ? subtotalMatch[1] : '';

  // name: remove JAN, qty, subtotal (best-effort)
  let name = line;
  if (jan) name = name.replace(jan, '').trim();
  name = name.replace(/\(\d+\s*個\)/, '').trim();
  name = name.replace(/小計\s*[0-9,]+円/, '').trim();

  // sometimes there are extra separators
  name = name.replace(/^[-:：\s]+/, '').trim();

  // final cleaning per your request
  name = cleanProductName(name);

  return { jan, name, qty, subtotal, raw: line };
}

async function authorize() {
  const { installed, web } = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
  const key = installed || web;
  const { client_id, client_secret, redirect_uris } = key;

  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris?.[0]
  );

  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf-8'));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });

  console.log('Authorize this app by visiting this url:', authUrl);

  const rl = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const code = await new Promise(resolve =>
    rl.question('Enter the code from that page here: ', c => {
      rl.close();
      resolve(c);
    })
  );

  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);

  fs.writeFileSync(
    TOKEN_PATH,
    JSON.stringify(
      {
        type: 'authorized_user',
        client_id,
        client_secret,
        refresh_token: tokens.refresh_token,
      },
      null,
      2
    )
  );

  return oAuth2Client;
}

// ======================================================
// List & export Pokemon Center "Order Completed" mails (NON-UNIQUE)
// ======================================================
async function listPokemonOrderComplete(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

  const SUBJECT = '[ポケモンセンターオンライン]注文完了のお知らせ';

  const res = await gmail.users.messages.list({
    userId: 'me',
    q: `subject:"${SUBJECT}" newer_than:7d`,
    maxResults: 500,
  });

  const messages = res.data.messages || [];
  if (messages.length === 0) {
    console.log(`Không tìm thấy email subject: ${SUBJECT}`);
    return;
  }

  console.log('=====================');
  console.log(`Found messages: ${messages.length}`);
  console.log('=====================');

  // Excel rows (one row per product line)
  const rows = [];

  // Log list (message-level)
  const mailList = [];

  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'full',
    });

    const headers = msg.data.payload?.headers || [];
    const subject = getHeader(headers, 'Subject').trim();
    const fromHeader = getHeader(headers, 'From').trim();
    const toHeader = getHeader(headers, 'To').trim();
    const dateHeader = getHeader(headers, 'Date').trim();

    // Decide "To" column email (only email)
    const toEmail = decideToEmailForRow(fromHeader, toHeader);

    // Collect body text
    const payload = msg.data.payload;
    const bodies = collectBodyTexts(payload);
    const plain = bodies.plain.join('\n').trim();
    const html = bodies.html.join('\n').trim();
    const fullText = plain || htmlToText(html);

    // Extract product lines
    const productLines = extractProductLines(fullText);

    // Mail log summary (your example style)
    mailList.push({
      date: dateHeader,
      toEmail,
      from: fromHeader,
      items: productLines.length,
    });

    // Add one row per product line
    if (productLines.length === 0) {
      rows.push({
        date: dateHeader,
        to: toEmail,
        jan: '',
        name: '',
        qty: '',
        subtotal: '',
      });
      continue;
    }

    for (const line of productLines) {
      const p = parseProductLine(line);
      rows.push({
        date: dateHeader,
        to: toEmail,
        jan: p.jan,
        name: p.name,
        qty: p.qty,
        subtotal: p.subtotal,
      });
    }
  }

  // ===================== LOGS =====================
  console.log('\n========== MAIL LIST ==========');
  console.log(`Messages: ${mailList.length}`);
  for (const it of mailList) {
    console.log(
      `date=${it.date || 'N/A'} | to=${it.toEmail || 'N/A'} | from=${it.from || 'N/A'} | items=${it.items}`
    );
  }
  console.log('===============================');

  console.log('\n========== EXPORT STATS ==========');
  console.log(`Excel rows: ${rows.length}`);
  console.log('=================================');

  // ======================================================
  // EXPORT EXCEL (ONLY requested columns)
  // ======================================================
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'gmail-order-complete-export';
  workbook.created = new Date();

  const ws = workbook.addWorksheet('Orders');
  ws.columns = [
    { header: 'Date', key: 'date', width: 28 },
    { header: 'To', key: 'to', width: 30 },
    { header: 'JAN', key: 'jan', width: 16 },
    { header: 'Product Name', key: 'name', width: 70 },
    { header: 'Qty', key: 'qty', width: 8 },
    { header: 'Subtotal', key: 'subtotal', width: 12 },
  ];
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.getRow(1).font = { bold: true };
  ws.autoFilter = { from: 'A1', to: 'F1' };

  rows.forEach(r => ws.addRow(r));

  const outPath = path.join(__dirname, 'gmail_order_complete.xlsx');
  await workbook.xlsx.writeFile(outPath);

  console.log(`\nExcel exported: ${outPath}`);
}

// Run
authorize()
  .then(auth => listPokemonOrderComplete(auth))
  .catch(console.error);
