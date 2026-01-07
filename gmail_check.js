const { google } = require('googleapis');
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
    q: 'subject:当選 OR subject:抽選結果 newer_than:30d',
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
  // ✅ Direction B summary: UNIQUE counts (consistent with CSV)
  // ======================================================
  const winUnique = [...resultMap.values()].filter(v => v === 'o').length;
  const loseUnique = [...resultMap.values()].filter(v => v === 'x').length;
  const totalUnique = resultMap.size;

  console.log('=====================');
  console.log(`（当選 unique: ${winUnique}）`);
  console.log(`（落選 unique: ${loseUnique}）`);
  console.log(`Unique total（当選＋落選）: ${totalUnique}`);
  console.log('=====================');

  // ======================================================
  // EXPORT CSV — 当選(o) first → 落選(x) after (UNIQUE)
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

  const lines = ['Email;Result'];
  winList.forEach(r => lines.push(`${r.mail};${r.result}`));
  loseList.forEach(r => lines.push(`${r.mail};${r.result}`));

  const outPath = path.join(__dirname, 'gmail_lottery_result.csv');
  fs.writeFileSync(outPath, lines.join('\n'), 'utf8');

  console.log(`CSV exported: ${outPath}`);
}

// Run
authorize()
  .then(auth => listPokemonLottery(auth))
  .catch(console.error);
