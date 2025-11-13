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

  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);
  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf-8'));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = require('readline').createInterface({ input: process.stdin, output: process.stdout });
  const code = await new Promise(resolve => rl.question('Enter the code from that page here: ', c => { rl.close(); resolve(c); }));
  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);

  fs.writeFileSync(TOKEN_PATH, JSON.stringify({
    type: 'authorized_user',
    client_id, client_secret,
    refresh_token: tokens.refresh_token,
  }));
  return oAuth2Client;
}

// 🧾 List & thống kê mail [ポケモンセンターオンライン]注文完了のお知らせ
async function listPokemonLottery(auth) {
  const gmail = google.gmail({ version: 'v1', auth });

  const res = await gmail.users.messages.list({
    userId: 'me',
    q: 'subject:[ポケモンセンターオンライン]注文完了のお知らせ newer_than:30d',
    maxResults: 500,
  });

  const messages = res.data.messages || [];
  if (messages.length === 0) {
    console.log('Không tìm thấy email [ポケモンセンターオンライン]注文完了のお知らせ.');
    return;
  }

  const purchasedMails = [];
  // const loseMails = [];

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
    const to = getHeader(headers, 'To');

    if (subject.includes('[ポケモンセンターオンライン]注文完了のお知らせ')) {
      purchasedMails.push({ from, to });
    }
  }
  // 🟩 In danh sách 注文完了 trước
  console.log('===== 🎉 注文完了 =====');
  purchasedMails.forEach(m => {
    console.log(`注文完了: ${m.from} | To: ${m.to}`);
  });
  console.log(`注文完了: ${purchasedMails.length}）\n`);
}

authorize()
  .then(auth => listPokemonLottery(auth))
  .catch(console.error);


  // 当選: paring.tweed
  // 落選: minaret_razz.9