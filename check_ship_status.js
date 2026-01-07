const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
const TOKEN_PATH = path.join(__dirname, 'token.json'); // if you still use token.json
const CREDENTIALS_PATH = path.join(__dirname, 'pokemon_cre.json');

// ------------------------------------------------------
// Helper: get header value (Subject, From, To, etc.)
// ------------------------------------------------------
function getHeader(headers, name) {
  const h = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
  return h ? h.value : '';
}

// ------------------------------------------------------
// Helper: decode Gmail base64url body
// ------------------------------------------------------
function decodeBase64Url(data) {
  return Buffer.from(
    data.replace(/-/g, '+').replace(/_/g, '/'),
    'base64'
  ).toString('utf-8');
}

// ------------------------------------------------------
// Helper: recursively get body text from payload
// ------------------------------------------------------
function getBodyFromPayload(payload) {
  // If this part has data directly
  if (payload.body && payload.body.data) {
    return decodeBase64Url(payload.body.data);
  }

  // If it has sub parts, search them
  if (payload.parts && payload.parts.length) {
    for (const part of payload.parts) {
      const text = getBodyFromPayload(part);
      if (text) return text; // Return first non-empty
    }
  }

  return '';
}

// ------------------------------------------------------
// Helper: extract WaybillNo (送り状番号 / お問い合わせ伝票番号)
// from body text or tracking URL
// ------------------------------------------------------
function extractWaybillNo(bodyText, trackingUrl) {
  if (!bodyText) bodyText = '';

  // 1) Simple pattern: label + : + number (allow spaces)
  let m = bodyText.match(
    /(送り状番号|お問い合わせ伝票番号)[：:]\s*([0-9\-]{5,})/
  );
  if (m) return m[2];

  // 2) Allow HTML or other chars between label and number
  //    e.g. 送り状番号：</th><td>123456789012</td>
  m = bodyText.match(
    /(送り状番号|お問い合わせ伝票番号)[^0-9]{0,50}([0-9\-]{5,})/
  );
  if (m) return m[2];

  // 3) Allow line break after label
  //    e.g.
  //       送り状番号：
  //       123456789012
  m = bodyText.match(
    /(送り状番号|お問い合わせ伝票番号)[^0-9\r\n]{0,10}[\r\n]+[^\r\n]*?([0-9\-]{5,})/
  );
  if (m) return m[2];

  // 4) Fallback: extract from tracking URL (if numeric pno)
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

  // Normalize line endings and split to lines
  const lines = bodyText.replace(/\r\n/g, '\n').split('\n').map(l => l.trim());

  let idx = lines.findIndex(line => line.includes('お届け先'));
  if (idx === -1) {
    return '';
  }

  let name = '';
  let zip = '';
  const addressLines = [];

  // Scan a few lines after お届け先
  for (let i = idx + 1; i < lines.length; i++) {
    const line = lines[i].trim();

    // Stop if we reach another section
    if (
      line.startsWith('【') || // e.g. 【ご注文者】
      line.includes('お支払い方法') ||
      line.includes('ご注文商品') ||
      line.includes('ご注文内容')
    ) {
      break;
    }

    if (!line) continue;

    // Name line (with 様)
    if (!name && line.includes('様')) {
      name = line;
      continue;
    }

    // Zip code line (starts with 〒)
    if (!zip && line.startsWith('〒')) {
      zip = line;
      continue;
    }

    // Address lines
    addressLines.push(line);
  }

  // Build final address string
  const parts = [];
  if (name) parts.push(name);
  if (zip) parts.push(zip);
  if (addressLines.length > 0) parts.push(addressLines.join(' '));

  return parts.join(' ');
}

// ------------------------------------------------------
// OAuth2 authorize
// ------------------------------------------------------
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
    JSON.stringify({
      type: 'authorized_user',
      client_id,
      client_secret,
      refresh_token: tokens.refresh_token,
    })
  );
  return oAuth2Client;
}

// ======================================================
// 📦 List & export shipping mails (出荷メール)
// ======================================================
async function listPokemonShippingMails(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

  // Search shipping mails from Pokemon Center Online (last 30 days)
  const res = await gmail.users.messages.list({
    userId: 'me',
    q: 'subject:"【ポケモンセンターオンライン】商品が出荷されました" newer_than:30d',
    maxResults: 500,
  });

  const messages = res.data.messages || [];

  if (messages.length === 0) {
    console.log('「【ポケモンセンターオンライン】商品が出荷されました」メールが見つかりません。');
    return;
  }

  const records = [];

  for (const m of messages) {
    // Get full message to read body
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'full',
    });

    const payload = msg.data.payload;
    const headers = payload.headers || [];

    const subject = getHeader(headers, 'Subject').trim();
    const from = getHeader(headers, 'From');
    const toHeader = getHeader(headers, 'To') || '';

    // Get body text
    const bodyText = getBodyFromPayload(payload) || '';

    // Extract Kuroneko tracking URL(s)
    const urlRegex = /https:\/\/member\.kms\.kuronekoyamato\.co\.jp\/parcel\/detail\?pno=[A-Za-z0-9]+/g;
    const urlMatches = bodyText.match(urlRegex) || [];
    const trackingUrl = urlMatches.length > 0 ? urlMatches[0] : '';

    // Extract Waybill number (送り状番号 / お問い合わせ伝票番号)
    const waybillNo = extractWaybillNo(bodyText, trackingUrl);

    // Extract shipping address block after お届け先
    const address = extractAddress(bodyText);

    // Parse To header to get each email address
    const toList = toHeader
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    for (const addr of toList) {
      const match = addr.match(/<([^>]+)>/);
      const email = match ? match[1] : addr;
      if (!email) continue;

      records.push({
        email,
        waybillNo,
        trackingUrl,
        address,
        subject,
        from,
      });
    }
  }

  console.log('===== 出荷メール一覧 =====');
  records.forEach(r => {
    console.log(
      `Email: ${r.email} | 送り状番号: ${r.waybillNo} | 住所: ${r.address} | URL: ${r.trackingUrl}`
    );
  });
  console.log('総件数:', records.length);
  console.log('=========================');

  // ======================================================
  // 📌 EXPORT CSV — Email;WaybillNo;TrackingUrl;Address
  // ======================================================
  const lines = ['Email;WaybillNo;TrackingUrl;Address'];

  records.forEach(r => {
    const email = (r.email || '').replace(/;/g, ',');
    const waybillNo = (r.waybillNo || '').replace(/;/g, ',');
    const trackingUrl = (r.trackingUrl || '').replace(/;/g, ',');
    const address = (r.address || '').replace(/;/g, ',');

    lines.push(`${email};${waybillNo};${trackingUrl};${address}`);
  });

  const csvContent = lines.join('\n');
  const outPath = path.join(__dirname, 'gmail_pokemon_shipping.csv');

  fs.writeFileSync(outPath, csvContent, 'utf8');
  console.log(`CSV exported: ${outPath}`);
}

// ======================================================

authorize()
  .then(auth => listPokemonShippingMails(auth))
  .catch(console.error);
