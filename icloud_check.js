// iCloud IMAP: export lottery result to CSV (Direction B: unique "To" mailbox results)
// Result format: mail,result  (o = win / x = lose)
//
// Notes:
// - Console counts are UNIQUE mailbox counts, consistent with CSV.
// - If an address has both win/lose emails, win ("o") wins and is not overwritten.
//
// SECURITY:
// - Do NOT hardcode iCloud credentials. Use environment variables (.env).
// - Revoke the leaked app-specific password immediately.

const { ImapFlow } = require('imapflow');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

function isWinSubject(subject) {
  // Adjust as needed
  return subject.includes('【新商品】2025年11月12日号') || subject.includes('当選');
}

function isLoseSubject(subject) {
  // Adjust as needed
  return subject.includes('抽選結果');
}

async function listIcloudLottery() {
  const ICLOUD_USER = 'phanhangnga2001@icloud.com';
  const ICLOUD_PASS = 'kdkv-dxtj-drtk-thpo';
  // const ICLOUD_PASS = process.env.ICLOUD_PASS;
  // const ICLOUD_PASS = process.env.ICLOUD_PASS;

  if (!ICLOUD_USER || !ICLOUD_PASS) {
    console.error('Missing ICLOUD_USER / ICLOUD_PASS in .env');
    process.exit(1);
  }

  const client = new ImapFlow({
    host: 'imap.mail.me.com',
    port: 993,
    secure: true,
    auth: {
      user: ICLOUD_USER, // Apple ID email
      pass: ICLOUD_PASS, // app-specific password
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
    sinceDate.setDate(sinceDate.getDate() - 7); // last 1 week

    const messageUids = await client.search({ since: sinceDate });

    // These arrays count MESSAGE occurrences (not unique). Keep if you still want raw message logs.
    const winMessages = [];
    const loseMessages = [];

    // key: email address (To), value: 'o' (win) or 'x' (lose)
    const resultMap = new Map();

    for await (const msg of client.fetch(messageUids, { envelope: true })) {
      const subject = msg.envelope.subject || '';
      const fromList = msg.envelope.from || [];
      const toList = msg.envelope.to || [];

      const from = fromList.map(a => `${a.name || ''} <${a.address}>`).join(', ');
      const to = toList.map(a => `${a.name || ''} <${a.address}>`).join(', ');

      const win = isWinSubject(subject);
      const lose = !win && isLoseSubject(subject); // avoid double count if subject matches both

      if (!win && !lose) continue; // skip unrelated mails

      // Keep message-level logs (optional)
      if (win) winMessages.push({ from, to, subject });
      if (lose) loseMessages.push({ from, to, subject });

      // Update result map per "To" email (unique mailbox)
      for (const addr of toList) {
        const mail = addr.address;
        if (!mail) continue;

        const current = resultMap.get(mail);
        if (win) {
          resultMap.set(mail, 'o'); // win always wins
        } else if (lose) {
          if (current !== 'o') resultMap.set(mail, 'x'); // do not overwrite win
        }
      }
    }

    // ======================================================
    // ✅ Direction B: UNIQUE counts (consistent with CSV)
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
    // 📌 EXPORT CSV — 当選(o) first → 落選(x) after (UNIQUE)
    // ======================================================
    const winList = [];
    const loseList = [];

    for (const [mail, result] of resultMap.entries()) {
      if (result === 'o') winList.push({ mail, result });
      else if (result === 'x') loseList.push({ mail, result });
    }

    // Sort optional (stable output)
    winList.sort((a, b) => a.mail.localeCompare(b.mail));
    loseList.sort((a, b) => a.mail.localeCompare(b.mail));

    const lines = ['mail,result'];
    winList.forEach(r => lines.push(`${r.mail},${r.result}`));
    loseList.forEach(r => lines.push(`${r.mail},${r.result}`));

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