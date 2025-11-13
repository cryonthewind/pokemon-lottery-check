// iCloud IMAP: export lottery result to CSV
// Result format: mail,result  (o = win / x = lose)

const { ImapFlow } = require('imapflow');
const fs = require('fs');
const path = require('path');

async function listIcloudLottery() {
  const client = new ImapFlow({
    host: 'imap.mail.me.com',
    port: 993,
    secure: true,
    auth: {
      user: 'iris_992000@icloud.com',   // Apple ID email
      pass: 'jgxg-iskc-zfhn-joeq',      // app-specific password
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
    sinceDate.setMonth(sinceDate.getMonth() - 1); // last 1 month

    const messageUids = await client.search({ since: sinceDate });

    const winMails = [];
    const loseMails = [];

    // Map to store final result per email
    // key: email address, value: 'o' (win) or 'x' (lose)
    const resultMap = new Map();

    for await (const msg of client.fetch(messageUids, { envelope: true })) {
      const subject = msg.envelope.subject || '';
      const fromList = msg.envelope.from || [];
      const toList = msg.envelope.to || [];

      const from = fromList.map(a => `${a.name || ''} <${a.address}>`).join(', ');
      const to = toList.map(a => `${a.name || ''} <${a.address}>`).join(', ');

      // ====== Detect win / lose by subject ======
      let isWin = false;
      let isLose = false;

      // You can adjust these conditions as you like
      if (subject.includes('【新商品】2025年11月12日号') || subject.includes('当選')) {
        isWin = true;
      } else if (subject.includes('抽選結果')) {
        isLose = true;
      }

      if (!isWin && !isLose) continue; // skip unrelated mails

      // ====== Console list (same as before) ======
      if (isWin) {
        winMails.push({ from, to });
      } else if (isLose) {
        loseMails.push({ from, to });
      }

      // ====== Update result map per "To" email ======
      for (const addr of toList) {
        const mail = addr.address;
        if (!mail) continue;

        // If already marked as win, do not overwrite with lose
        const current = resultMap.get(mail);
        if (isWin) {
          resultMap.set(mail, 'o'); // win
        } else if (isLose) {
          if (current !== 'o') {
            resultMap.set(mail, 'x'); // lose
          }
        }
      }
    }

    // ====== Print to console (optional) ======
    console.log('===== 🎉 当選 =====');
    winMails.forEach(m => {
      console.log(`当選、From: ${m.from} | To: ${m.to}`);
    });
    console.log(`（当選: ${winMails.length}）\n`);

    console.log('===== 💧 落選（抽選結果） =====');
    loseMails.forEach(m => {
      console.log(`落選、From: ${m.from} | To: ${m.to}`);
    });
    console.log(`（落選: ${loseMails.length}）\n`);

    const total = winMails.length + loseMails.length;
    console.log('=====================');
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
    const outPath = path.join(__dirname, 'icloud_lottery_result.csv');
  
    fs.writeFileSync(outPath, csvContent, 'utf8');

    console.log(`CSV exported: ${outPath}`);
  } catch (err) {
    console.error(err);
  } finally {
    if (!client.closed) {
      await client.logout().catch(() => {});
    }
  }
}

listIcloudLottery();
