// Gmail: export Pokemon Center lottery result to EXCEL (Direction B: unique mailbox results)
// Result format: Email, Result (o = win / x = lose)
//
// Notes:
// - Console counts are UNIQUE mailbox counts, consistent with Excel.
// - If an address has both win/lose emails, win ("o") wins and is not overwritten.

const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'pokemon_cre.json');

function getHeader(headers, name) {
  const h = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
  return h ? h.value : '';
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

// Decide which header to use for mapping result:
// - If sender is hotmail/outlook => use FROM
// - Otherwise => use TO (Pokemon Center pattern)
function decideTargetEmails(from, toHeader) {
  if (/hotmail\.com|outlook\.com/i.test(from || '')) {
    return extractEmails(from);
  }
  return extractEmails(toHeader);
}

// ===================== log helpers =====================
function pct(n, d) {
  if (!d) return '0.00%';
  return `${((n / d) * 100).toFixed(2)}%`;
}
// ======================================================

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
// List & export lottery mails 当選 / 落選 (Direction B: UNIQUE mailbox results)
// ======================================================
async function listPokemonLottery(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

  // Query: last 30 days, subjects containing 当選 or 抽選結果
  const res = await gmail.users.messages.list({
    userId: 'me',
    q: 'subject:当選 OR subject:抽選結果 newer_than:1d',
    maxResults: 500,
  });

  const messages = res.data.messages || [];
  if (messages.length === 0) {
    console.log('Không tìm thấy email 当選 hoặc 抽選結果.');
    return;
  }

  // Message-level logs (these count messages, NOT unique mailboxes)
  const winMails = [];
  const loseMails = [];

  // Unique mailbox result map: email -> 'o' (win) or 'x' (lose)
  const resultMap = new Map();

  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'metadata',
      metadataHeaders: ['Subject', 'From', 'To'],
    });

    const headers = msg.data.payload?.headers || [];
    const subject = getHeader(headers, 'Subject').trim();
    const from = getHeader(headers, 'From').trim();
    const toHeader = getHeader(headers, 'To').trim();

    // Detect win/lose by subject
    const isWin = subject.includes('当選');
    const isLose = !isWin && subject.includes('抽選結果'); // avoid double count if both match

    if (!isWin && !isLose) continue;

    // Keep message-level logs (optional)
    if (isWin) winMails.push({ from, to: toHeader, subject });
    if (isLose) loseMails.push({ from, to: toHeader, subject });

    // Choose mapping target (FROM for hotmail/outlook, else TO)
    const targetEmails = decideTargetEmails(from, toHeader);
    if (targetEmails.length === 0) continue;

    // Update unique result map (win overrides lose)
    for (const email of targetEmails) {
      const current = resultMap.get(email);
      if (isWin) {
        resultMap.set(email, 'o');
      } else if (isLose) {
        if (current !== 'o') resultMap.set(email, 'x');
      }
    }
  }

  // ======================================================
  // ✅ Direction B summary: UNIQUE counts (consistent with Excel)
  // ======================================================
  const winUnique = [...resultMap.values()].filter(v => v === 'o').length;
  const loseUnique = [...resultMap.values()].filter(v => v === 'x').length;
  const totalUnique = resultMap.size;

  console.log('=====================');
  console.log(`（当選 unique: ${winUnique}）`);
  console.log(`（落選 unique: ${loseUnique}）`);
  console.log(`Unique total（当選＋落選）: ${totalUnique}`);
  console.log('=====================');

  // ===================== ONLY final logs =====================
  // Build unique email lists from resultMap (no logic change)
  const winEmails = [];
  const loseEmails = [];

  for (const [mail, result] of resultMap.entries()) {
    if (result === 'o') winEmails.push(mail);
    else if (result === 'x') loseEmails.push(mail);
  }

  // Stable sort for readability
  winEmails.sort((a, b) => a.localeCompare(b));
  loseEmails.sort((a, b) => a.localeCompare(b));

  // Rate log
  console.log('\n========== HIT RATE ==========');
  console.log(`当選率: ${pct(winUnique, totalUnique)} (${winUnique}/${totalUnique})`);
  console.log(`落選率: ${pct(loseUnique, totalUnique)} (${loseUnique}/${totalUnique})`);
  console.log('==============================');

  // Detail log (your requested format)
  console.log('\n========== CHECK DETAIL ==========');

  console.log(`当選 emails (${winEmails.length}):`);
  for (const mail of winEmails) {
    console.log('  +', mail);
  }

  console.log(`\n落選 emails (${loseEmails.length}):`);
  for (const mail of loseEmails) {
    console.log('  -', mail);
  }

  console.log('==================================');
  // =================== END LOGS ===================

  // ======================================================
  // EXPORT EXCEL — 当選(o) first → 落選(x) after (UNIQUE)
  // ======================================================
  const winList = [];
  const loseList = [];

  for (const [mail, result] of resultMap.entries()) {
    if (result === 'o') winList.push({ mail, result });
    else if (result === 'x') loseList.push({ mail, result });
  }

  // Stable sort
  winList.sort((a, b) => a.mail.localeCompare(b.mail));
  loseList.sort((a, b) => a.mail.localeCompare(b.mail));

  // Create Excel workbook
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'gmail-lottery-export';
  workbook.created = new Date();

  // Sheet 1: Results (o first, then x)
  const ws = workbook.addWorksheet('Results');
  ws.columns = [
    { header: 'Email', key: 'mail', width: 40 },
    { header: 'Result', key: 'result', width: 10 },
  ];
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.getRow(1).font = { bold: true };
  ws.autoFilter = { from: 'A1', to: 'B1' };

  // Add rows (o first then x)
  winList.forEach(r => ws.addRow(r));
  loseList.forEach(r => ws.addRow(r));

  // Sheet 2: Summary (optional)
  const ws2 = workbook.addWorksheet('Summary');
  ws2.columns = [
    { header: 'Metric', key: 'metric', width: 25 },
    { header: 'Value', key: 'value', width: 35 },
  ];
  ws2.getRow(1).font = { bold: true };
  ws2.addRow({ metric: 'Win (unique)', value: winUnique });
  ws2.addRow({ metric: 'Lose (unique)', value: loseUnique });
  ws2.addRow({ metric: 'Total (unique)', value: totalUnique });
  ws2.addRow({ metric: 'Win rate', value: `${pct(winUnique, totalUnique)} (${winUnique}/${totalUnique})` });
  ws2.addRow({ metric: 'Lose rate', value: `${pct(loseUnique, totalUnique)} (${loseUnique}/${totalUnique})` });

  const outPath = path.join(__dirname, 'gmail_lottery_result.xlsx');
  await workbook.xlsx.writeFile(outPath);

  console.log(`Excel exported: ${outPath}`);
}

// Run
authorize()
  .then(auth => listPokemonLottery(auth))
  .catch(console.error);
