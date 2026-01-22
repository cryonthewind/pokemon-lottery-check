// NOTE: Code comments are in English as requested.

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'pokemon_cre.json');

// ------------------------------------------------------
// Helper: get header value (Subject, From, To, etc.)
// ------------------------------------------------------
function getHeader(headers, name) {
  const h = (headers || []).find(h => (h.name || '').toLowerCase() === name.toLowerCase());
  return h ? (h.value || '') : '';
}

// ------------------------------------------------------
// Helper: extract emails from header like:
// 'Name <a@b.com>, "X" <c@d.com>' => ['a@b.com','c@d.com']
// 'a@b.com' => ['a@b.com']
// ------------------------------------------------------
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

// ------------------------------------------------------
// Helper: decode Gmail base64url body
// ------------------------------------------------------
function decodeBase64Url(data) {
  if (!data) return '';
  const b64 = data.replace(/-/g, '+').replace(/_/g, '/');
  const pad = b64.length % 4 ? '='.repeat(4 - (b64.length % 4)) : '';
  return Buffer.from(b64 + pad, 'base64').toString('utf-8');
}

// ------------------------------------------------------
// Helper: walk payload and collect text/plain and text/html
// ------------------------------------------------------
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

// ------------------------------------------------------
// Helper: light HTML -> text fallback
// ------------------------------------------------------
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

// ------------------------------------------------------
// Helper: extract WaybillNo (送り状番号 / お問い合わせ伝票番号)
// from body text or tracking URL
// ------------------------------------------------------
function extractWaybillNo(bodyText, trackingUrl) {
  if (!bodyText) bodyText = '';

  let m = bodyText.match(/(送り状番号|お問い合わせ伝票番号)[：:]\s*([0-9\-]{5,})/);
  if (m) return m[2];

  m = bodyText.match(/(送り状番号|お問い合わせ伝票番号)[^0-9]{0,50}([0-9\-]{5,})/);
  if (m) return m[2];

  m = bodyText.match(
    /(送り状番号|お問い合わせ伝票番号)[^0-9\r\n]{0,10}[\r\n]+[^\r\n]*?([0-9\-]{5,})/
  );
  if (m) return m[2];

  if (trackingUrl) {
    const urlMatch = trackingUrl.match(/pno=([0-9\-]{5,})/);
    if (urlMatch) return urlMatch[1];
  }

  return '';
}

// ------------------------------------------------------
// Helper: extract shipping address block after お届け先
// ------------------------------------------------------
function extractAddress(bodyText) {
  if (!bodyText) return '';

  const lines = bodyText.replace(/\r\n/g, '\n').split('\n').map(l => (l || '').trim());
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
// Clean product name:
// - remove leading 【抽選販売】
// - normalize spaces
// ------------------------------------------------------
function cleanProductName(name) {
  if (!name) return '';
  let s = name.trim();

  s = s.replace(/^【抽選販売】\s*/g, '');
  s = s.replace(/\s+/g, ' ').trim();

  return s;
}

// ------------------------------------------------------
// Extract product block lines under 【商品情報】
// IMPORTANT: Shipping mail may NOT include "小計".
// Example line:
// 9900000007003 【抽選販売】ポケモンカードゲーム ... BOX 5,400円 1個
// ------------------------------------------------------
function extractProductLinesShipping(fullText) {
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
    '【お届け先】',
    'お届け先',
    '配送情報',
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
    .map(l => (l || '').trim())
    .filter(Boolean);

  // Keep only lines likely to be product lines:
  // - have price "円"
  // - and qty "個"
  // - and have some text
  return lines.filter(l => l.includes('円') && l.includes('個'));
}

// ------------------------------------------------------
// Parse one shipping product line into { name, price }
// Example:
// 9900000007003 【抽選販売】ポケモンカードゲーム ... BOX 5,400円 1個
// ------------------------------------------------------
function parseShippingProductLine(line) {
  const janMatch = line.match(/\b(\d{8,14})\b/);
  const qtyMatch = line.match(/(\d+)\s*個/);
  const priceMatch = line.match(/([0-9,]+)\s*円/);

  const jan = janMatch ? janMatch[1] : '';
  const qty = qtyMatch ? qtyMatch[1] : '';
  const price = priceMatch ? priceMatch[1] : ''; // keep "5,400" (no 円)

  // Remove known tokens to isolate name
  let name = line;

  if (jan) name = name.replace(jan, '').trim();
  // Remove price + 円
  if (price) name = name.replace(new RegExp(`${price}\\s*円`), '').trim();
  // Remove qty + 個
  if (qty) name = name.replace(new RegExp(`${qty}\\s*個`), '').trim();

  // Clean leading separators
  name = name.replace(/^[-:：\s]+/, '').trim();

  // Final cleaning (remove 【抽選販売】, normalize spaces)
  name = cleanProductName(name);

  return { name, price, raw: line };
}

// ------------------------------------------------------
// Extract Product Name + Price for SHIPPING mail
// If multiple products: join by " / "
// ------------------------------------------------------
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

  // De-duplicate while preserving order
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

  const uniqNames = uniq(names);
  const uniqPrices = uniq(prices);

  return {
    productName: uniqNames.join(' / '),
    price: uniqPrices.join(' / '),
  };
}

// ------------------------------------------------------
// OAuth2 authorize
// ------------------------------------------------------
async function authorize() {
  const { installed, web } = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
  const key = installed || web;
  const { client_id, client_secret, redirect_uris } = key;

  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);

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

// ------------------------------------------------------
// Export records to Excel (.xlsx)
// IMPORTANT: Replace From column by Price column
// Columns: Email, WaybillNo, TrackingUrl, Address, Product Name, Price
// ------------------------------------------------------
async function exportToExcel(records) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('ShippingMails');

  sheet.columns = [
    { header: 'Email', key: 'email', width: 35 },
    { header: 'WaybillNo', key: 'waybillNo', width: 18 },
    { header: 'TrackingUrl', key: 'trackingUrl', width: 55 },
    { header: 'Address', key: 'address', width: 55 },
    { header: 'Product Name', key: 'productName', width: 70 },
    { header: 'Price', key: 'price', width: 12 }, // ✅ added, replaces From
  ];

  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { vertical: 'middle' };

  for (const r of records) {
    sheet.addRow({
      email: r.email || '',
      waybillNo: r.waybillNo || '',
      trackingUrl: r.trackingUrl || '',
      address: r.address || '',
      productName: r.productName || '',
      price: r.price || '',
    });
  }

  // Make URL clickable
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const cell = row.getCell('trackingUrl');
    const url = cell.value;
    if (typeof url === 'string' && url.startsWith('http')) {
      cell.value = { text: url, hyperlink: url };
    }
  });

  // Enable wrap text
  sheet.getColumn('address').alignment = { wrapText: true, vertical: 'top' };
  sheet.getColumn('productName').alignment = { wrapText: true, vertical: 'top' };
  sheet.getColumn('trackingUrl').alignment = { wrapText: true, vertical: 'top' };

  sheet.autoFilter = { from: 'A1', to: 'F1' };

  const outPath = path.join(__dirname, 'gmail_pokemon_shipping.xlsx');
  await workbook.xlsx.writeFile(outPath);

  console.log(`Excel exported: ${outPath}`);
}

// ======================================================
// 📦 List & export shipping mails (出荷メール)
// ======================================================
async function listPokemonShippingMails(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

  const res = await gmail.users.messages.list({
    userId: 'me',
    q: 'subject:"【ポケモンセンターオンライン】商品が出荷されました" newer_than:1d',
    maxResults: 500,
  });

  const messages = res.data.messages || [];
  if (messages.length === 0) {
    console.log('Không tìm thấy mail 「【ポケモンセンターオンライン】商品が出荷されました」');
    return;
  }

  const records = [];

  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'full',
    });

    const payload = msg.data.payload;
    const headers = payload?.headers || [];

    const toHeader = getHeader(headers, 'To') || '';

    // Collect body text (plain preferred)
    const bodies = collectBodyTexts(payload);
    const plain = bodies.plain.join('\n').trim();
    const html = bodies.html.join('\n').trim();
    const fullText = plain || htmlToText(html) || '';

    // Extract tracking URL (Kuroneko)
    const urlRegex = /https:\/\/member\.kms\.kuronekoyamato\.co\.jp\/parcel\/detail\?pno=[A-Za-z0-9]+/g;
    const urlMatches = fullText.match(urlRegex) || [];
    const trackingUrl = urlMatches.length > 0 ? urlMatches[0] : '';

    // Extract waybill number
    const waybillNo = extractWaybillNo(fullText, trackingUrl);

    // Extract address
    const address = extractAddress(fullText);

    // ✅ Extract product name + price (NEW)
    const { productName, price } = extractProductAndPriceForShipping(fullText);

    // Expand all "To" emails
    const toList = extractEmails(toHeader);
    for (const email of toList) {
      if (!email) continue;

      records.push({
        email,
        waybillNo,
        trackingUrl,
        address,
        productName, // ✅ now filled
        price,       // ✅ new column
      });
    }
  }

  console.log('===== SHIPPING MAIL LIST =====');
  for (const r of records) {
    console.log(
      `Email: ${r.email} | Waybill: ${r.waybillNo} | Product: ${r.productName} | Price: ${r.price} | Address: ${r.address} | URL: ${r.trackingUrl}`
    );
  }
  console.log('Total:', records.length);
  console.log('=============================');

  await exportToExcel(records);
}

// Run
authorize()
  .then(auth => listPokemonShippingMails(auth))
  .catch(console.error);
