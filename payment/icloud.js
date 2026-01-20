// iCloud IMAP: export Pokemon Center "Order Completed" to EXCEL (NON-UNIQUE)
// Subject: [ポケモンセンターオンライン]注文完了のお知らせ
//
// Excel columns (ONLY):
// Date, To, JAN, Product Name, Qty, Subtotal
//
// Rules:
// - No unique aggregation (same mailbox can appear many times)
// - If From is hotmail/outlook => use From email as "To"
// - Otherwise => use To email (first address)
// - "To" column must be pure email only (no display name)
// - Parse 【商品情報】 lines containing "小計" into rows (one row per product line)
// - Product Name: remove leading 【抽選販売】 and trailing 【...発送予定】 (best-effort)
//
// SECURITY:
// - Do NOT hardcode iCloud credentials. Use environment variables (.env).

const { ImapFlow } = require('imapflow');
const ExcelJS = require('exceljs');
const path = require('path');
require('dotenv').config();

// -------------------- Config --------------------
const SUBJECT_KEY = '[ポケモンセンターオンライン]注文完了のお知らせ';
const DAYS_BACK = Number(process.env.DAYS_BACK || 7); // default last 7 days
// ------------------------------------------------

function pct(n, d) {
  if (!d) return '0.00%';
  return `${((n / d) * 100).toFixed(2)}%`;
}

// Decide which mailbox to store in "To" column:
// - If sender is hotmail/outlook => use FROM email
// - Otherwise => use TO email (first)
function decideToEmailForRow(fromAddr, toAddr) {
  const fromEmail = (fromAddr && fromAddr.address) ? String(fromAddr.address).trim() : '';
  const toEmail = (toAddr && toAddr.address) ? String(toAddr.address).trim() : '';

  const isHotmail = /hotmail\.com|outlook\.com/i.test(fromEmail);
  return isHotmail ? fromEmail : toEmail;
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

  return lines.filter(l => l.includes('小計'));
}

// Clean product name:
// - remove leading 【抽選販売】
// - remove trailing schedule bracket like 【2026年...発送予定】
// - normalize spaces
function cleanProductName(name) {
  if (!name) return '';
  let s = name.trim();

  s = s.replace(/^【抽選販売】\s*/g, '');

  // remove trailing 【...発送予定】 or similar
  s = s.replace(/【[^【】]*発送[^【】]*】\s*$/g, '');

  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// Parse one product line into fields (best-effort)
function parseProductLine(line) {
  const janMatch = line.match(/\b(\d{8,14})\b/);
  const qtyMatch = line.match(/\((\d+)\s*個\)/);
  const subtotalMatch = line.match(/小計\s*([0-9,]+円)/);

  const jan = janMatch ? janMatch[1] : '';
  const qty = qtyMatch ? qtyMatch[1] : '';
  const subtotal = subtotalMatch ? subtotalMatch[1] : '';

  let name = line;
  if (jan) name = name.replace(jan, '').trim();
  name = name.replace(/\(\d+\s*個\)/, '').trim();
  name = name.replace(/小計\s*[0-9,]+円/, '').trim();
  name = name.replace(/^[-:：\s]+/, '').trim();

  name = cleanProductName(name);

  return { jan, name, qty, subtotal, raw: line };
}

async function listIcloudOrderComplete() {
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
      user: ICLOUD_USER,
      pass: ICLOUD_PASS,
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
    sinceDate.setDate(sinceDate.getDate() - DAYS_BACK);

    const messageUids = await client.search({ since: sinceDate });

    console.log('=====================');
    console.log(`Search since: ${sinceDate.toISOString()}`);
    console.log(`Matched UIDs: ${messageUids.length}`);
    console.log('=====================');

    const rows = [];     // Excel rows (one row per product line)
    const mailLogs = []; // Message-level logs

    let matchedMessages = 0;
    let parsedProductLinesTotal = 0;
    let noProductFound = 0;

    // Fetch envelope + source (raw RFC822)
    // "source" can be Buffer/string depending on server
    for await (const msg of client.fetch(messageUids, { envelope: true, source: true })) {
      const subject = msg.envelope?.subject || '';

      if (!subject.includes(SUBJECT_KEY)) continue;
      matchedMessages++;

      const fromList = msg.envelope?.from || [];
      const toList = msg.envelope?.to || [];

      const fromAddr = fromList[0] || null;
      const toAddr = toList[0] || null;

      const fromEmail = fromAddr?.address ? String(fromAddr.address).trim() : '';
      const toEmailRaw = toAddr?.address ? String(toAddr.address).trim() : '';

      // Your rule: hotmail/outlook => use FROM as "To", else use TO
      const toEmail = decideToEmailForRow(fromAddr, toAddr);

      // Date: prefer envelope date if available
      const dateObj = msg.envelope?.date ? new Date(msg.envelope.date) : null;
      const dateStr = dateObj && !Number.isNaN(dateObj.getTime())
        ? dateObj.toUTCString()
        : '';

      // Parse body text from raw source (simple + robust enough for this use)
      // We avoid heavy MIME parsing libs to keep changes minimal.
      // This best-effort approach still works well if email contains the 商品情報 lines in plain text.
      const raw = msg.source ? (Buffer.isBuffer(msg.source) ? msg.source.toString('utf-8') : String(msg.source)) : '';
      const rawText = raw.replace(/\r\n/g, '\n');

      // Best-effort: take everything after first blank line as "body"
      const splitIdx = rawText.indexOf('\n\n');
      const bodyText = splitIdx !== -1 ? rawText.slice(splitIdx + 2) : rawText;

      const productLines = extractProductLines(bodyText);

      mailLogs.push({
        date: dateStr,
        to: toEmail || toEmailRaw || '',
        from: fromEmail,
        items: productLines.length,
      });

      if (productLines.length === 0) {
        noProductFound++;
        // still write one row for traceability
        rows.push({
          date: dateStr,
          to: toEmail || toEmailRaw || '',
          jan: '',
          name: '',
          qty: '',
          subtotal: '',
        });
        continue;
      }

      parsedProductLinesTotal += productLines.length;

      for (const line of productLines) {
        const p = parseProductLine(line);
        rows.push({
          date: dateStr,
          to: toEmail || toEmailRaw || '',
          jan: p.jan,
          name: p.name,
          qty: p.qty,
          subtotal: p.subtotal,
        });
      }
    }

    // ===================== LOGS =====================
    console.log('\n========== MAIL LIST ==========');
    console.log(`Matched messages: ${matchedMessages}`);
    for (const it of mailLogs) {
      console.log(
        `date=${it.date || 'N/A'} | to=${it.to || 'N/A'} | from=${it.from || 'N/A'} | items=${it.items}`
      );
    }
    console.log('===============================');

    console.log('\n========== EXPORT STATS ==========');
    console.log(`Excel rows: ${rows.length}`);
    console.log(`Total product lines parsed: ${parsedProductLinesTotal}`);
    console.log(`Rows with NO product lines: ${noProductFound}`);
    if (matchedMessages) {
      console.log(`Avg product lines per message: ${(parsedProductLinesTotal / matchedMessages).toFixed(2)}`);
    }
    console.log('=================================');

    // ======================================================
    // EXPORT EXCEL (ONLY requested columns)
    // ======================================================
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'icloud-order-complete-export';
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

    const outPath = path.join(__dirname, 'icloud_order_complete.xlsx');
    await workbook.xlsx.writeFile(outPath);

    console.log(`\nExcel exported: ${outPath}`);
    console.log('\n========== DONE ==========');
    console.log(`Hit rate (matched/UIDs): ${pct(matchedMessages, messageUids.length)} (${matchedMessages}/${messageUids.length})`);
    console.log('==========================');

  } catch (err) {
    console.error(err);
  } finally {
    if (!client.closed) {
      await client.logout().catch(() => {});
    }
  }
}

listIcloudOrderComplete();
