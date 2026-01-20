// iCloud IMAP: export lottery result to EXCEL (Direction B: unique "To" mailbox results)
// Result format: mail,result  (o = win / x = lose)
//
// Notes:
// - Console counts are UNIQUE mailbox counts, consistent with Excel.
// - If an address has both win/lose emails, win ("o") wins and is not overwritten.
//
// SECURITY:
// - Do NOT hardcode iCloud credentials. Use environment variables (.env).

const { ImapFlow } = require('imapflow');
const ExcelJS = require('exceljs');
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

// Log helper
function pct(n, d) {
  if (!d) return '0.00%';
  return `${((n / d) * 100).toFixed(2)}%`;
}

async function listIcloudLottery() {
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
    // ✅ UNIQUE counts (consistent with Excel)
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
    // ✅ ADDED: HIT RATE + UNIQUE email lists (win/lose)
    // ======================================================
    const winEmails = [];
    const loseEmails = [];

    for (const [mail, result] of resultMap.entries()) {
      if (result === 'o') winEmails.push(mail);
      else if (result === 'x') loseEmails.push(mail);
    }

    winEmails.sort((a, b) => a.localeCompare(b));
    loseEmails.sort((a, b) => a.localeCompare(b));

    console.log('\n========== HIT RATE ==========');
    console.log(`当選率: ${pct(winUnique, totalUnique)} (${winUnique}/${totalUnique})`);
    console.log(`落選率: ${pct(loseUnique, totalUnique)} (${loseUnique}/${totalUnique})`);
    console.log('==============================');

    console.log('\n========== CHECK DETAIL (UNIQUE) ==========');
    console.log(`当選 emails (${winEmails.length}):`);
    for (const mail of winEmails) console.log('  +', mail);

    console.log(`\n落選 emails (${loseEmails.length}):`);
    for (const mail of loseEmails) console.log('  -', mail);
    console.log('==========================================');

    // ======================================================
    // 📌 EXPORT EXCEL — 当選(o) first → 落選(x) after (UNIQUE)
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

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'icloud-lottery-export';
    workbook.created = new Date();

    // Sheet 1: Results
    const ws = workbook.addWorksheet('Results');
    ws.columns = [
      { header: 'mail', key: 'mail', width: 40 },
      { header: 'result', key: 'result', width: 10 },
    ];
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.getRow(1).font = { bold: true };
    ws.autoFilter = { from: 'A1', to: 'B1' };

    // o first, then x (same as your CSV export ordering)
    winList.forEach(r => ws.addRow(r));
    loseList.forEach(r => ws.addRow(r));

    // Sheet 2: Summary
    const ws2 = workbook.addWorksheet('Summary');
    ws2.columns = [
      { header: 'metric', key: 'metric', width: 25 },
      { header: 'value', key: 'value', width: 40 },
    ];
    ws2.getRow(1).font = { bold: true };
    ws2.addRow({ metric: 'win_unique', value: winUnique });
    ws2.addRow({ metric: 'lose_unique', value: loseUnique });
    ws2.addRow({ metric: 'total_unique', value: totalUnique });
    ws2.addRow({ metric: 'win_rate', value: `${pct(winUnique, totalUnique)} (${winUnique}/${totalUnique})` });
    ws2.addRow({ metric: 'lose_rate', value: `${pct(loseUnique, totalUnique)} (${loseUnique}/${totalUnique})` });

    // Sheet 3: WinEmails (unique list)
    const ws3 = workbook.addWorksheet('WinEmails');
    ws3.columns = [{ header: 'mail', key: 'mail', width: 40 }];
    ws3.getRow(1).font = { bold: true };
    winEmails.forEach(mail => ws3.addRow({ mail }));

    // Sheet 4: LoseEmails (unique list)
    const ws4 = workbook.addWorksheet('LoseEmails');
    ws4.columns = [{ header: 'mail', key: 'mail', width: 40 }];
    ws4.getRow(1).font = { bold: true };
    loseEmails.forEach(mail => ws4.addRow({ mail }));

    const outPath = path.join(__dirname, 'icloud_lottery_result.xlsx');
    await workbook.xlsx.writeFile(outPath);

    console.log(`Excel exported: ${outPath}`);
  } catch (err) {
    console.error(err);
  } finally {
    if (!client.closed) {
      await client.logout().catch(() => {});
    }
  }
}

listIcloudLottery();
