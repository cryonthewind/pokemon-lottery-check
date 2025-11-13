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
// 🧾 List & export lottery mails 当選 / 落選
// ======================================================
async function listPokemonLottery(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

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

  const winMails = [];
  const loseMails = [];

  // email -> "o" / "x"
  const resultMap = new Map();

  for (const m of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: m.id,
      format: 'metadata',
      metadataHeaders: ['Subject', 'From', 'To'],
    });

    const headers = msg.data.payload.headers;
    const subject = getHeader(headers, 'Subject').trim();
    const from = getHeader(headers, 'From');
    const toHeader = getHeader(headers, 'To');

    let isWin = false;
    let isLose = false;

    if (subject.includes('当選')) {
      isWin = true;
    } else if (subject.includes('抽選結果')) {
      isLose = true;
    }

    if (!isWin && !isLose) continue;

    if (isWin) {
      winMails.push({ from, to: toHeader });
    } else if (isLose) {
      loseMails.push({ from, to: toHeader });
    }

    const toList = toHeader
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    for (const addr of toList) {
      const match = addr.match(/<([^>]+)>/);
      const email = match ? match[1] : addr;

      if (!email) continue;

      const current = resultMap.get(email);
      if (isWin) {
        resultMap.set(email, 'o');
      } else if (isLose) {
        if (current !== 'o') {
          resultMap.set(email, 'x');
        }
      }
    }
  }

  console.log('===== 🎉 当選 =====');
  winMails.forEach(m => {
    console.log(`当選、From: ${m.from} | To: ${m.to}`);
  });
  c

  console.log('===== 💧 落選 =====');
  loseMails.forEach(m => {
    console.log(`落選、From: ${m.from} | To: ${m.to}`);
  });
  console.log('=====================');
  onsole.log(`（当選: ${winMails.length}）\n`);
  console.log(`（落選: ${loseMails.length}）\n`);
  console.log(`抽選メール総数（当選＋落選）: ${total}`);
  console.log('=====================');
  // ======================================================
  // 📌 EXPORT CSV — 当選(o) trước → 落選(x) sau
  // ======================================================
  const winList = [];
  const loseList = [];

  for (const [mail, result] of resultMap.entries()) {
    if (result === 'o') winList.push({ mail, result });
    else if (result === 'x') loseList.push({ mail, result });
  }

  const lines = ['mail,result'];

  winList.forEach(r => lines.push(`${r.mail},${r.result}`)); // 当選 first
  loseList.forEach(r => lines.push(`${r.mail},${r.result}`)); // 落選 after

  const csvContent = lines.join('\n');
  const outPath = path.join(__dirname, 'gmail_lottery_result.csv');

  fs.writeFileSync(outPath, csvContent, 'utf8');
  console.log(`CSV exported: ${outPath}`);
}

// ======================================================

authorize()
  .then(auth => listPokemonLottery(auth))
  .catch(console.error);
// ======================================================