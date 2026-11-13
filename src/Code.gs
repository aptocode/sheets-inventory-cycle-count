/**
 * Daily Inventory Cycle Count for Google Sheets
 * ----------------------------------------------
 * Count a few items every day instead of the whole store once a month.
 *
 *  1. "Generate today's count" picks the items to count:
 *       - every High-risk item (the ones that get stolen or lost), plus
 *       - the items that have gone the longest without a count.
 *  2. Staff type the counted quantity in the "Count Today" sheet (phone or PC).
 *  3. "Submit count" saves the result to History, flags big differences
 *     and emails a short report to the manager.
 *
 * Sheets: Items · Count Today · History · Settings
 * License: MIT
 */

const SHEETS = {
  items: 'Items',
  count: 'Count Today',
  history: 'History',
  settings: 'Settings',
};

const ITEM_HEADERS = ['SKU', 'Item', 'Location', 'System Qty', 'Unit Cost', 'Risk', 'Last Counted'];
const COUNT_HEADERS = ['SKU', 'Item', 'Location', 'System Qty', 'Counted Qty', 'Difference', 'Difference $', 'Status'];
const HISTORY_HEADERS = ['Date', 'SKU', 'Item', 'Location', 'System Qty', 'Counted Qty', 'Difference', 'Difference $', 'Status', 'Counted By'];

const DEFAULT_SETTINGS = [
  ['Setting', 'Value', 'Notes'],
  ['Items per day', 15, 'Total items in the daily list (High-risk items are always included)'],
  ['Alert difference (units)', 2, 'Flag an item when the difference is this many units or more'],
  ['Alert difference ($)', 25, 'Flag an item when the difference is worth this much or more'],
  ['Report email', '', 'Leave empty to send the report to the sheet owner'],
  ['Update system qty after count', 'NO', 'YES = replace System Qty with the counted quantity'],
];

/* ------------------------------------------------------------------ menu */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cycle Count')
    .addItem("Generate today's count", 'generateDailyCount')
    .addItem('Submit count', 'submitCount')
    .addSeparator()
    .addItem('Run daily at 7 AM (auto)', 'installDailyTrigger')
    .addItem('Set up sheets with sample data', 'setupSheets')
    .addToUi();
}

/* ----------------------------------------------------------------- setup */

function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  const items = getOrCreate_(ss, SHEETS.items);
  if (items.getLastRow() < 2) {
    items.clear();
    items.appendRow(ITEM_HEADERS);
    SAMPLE_ITEMS.forEach(r => items.appendRow(r.concat([''])));
    items.getRange('E:E').setNumberFormat('$#,##0.00');
    items.getRange('G:G').setNumberFormat('yyyy-mm-dd');
  }

  const count = getOrCreate_(ss, SHEETS.count);
  if (count.getLastRow() === 0) count.appendRow(COUNT_HEADERS);

  const history = getOrCreate_(ss, SHEETS.history);
  if (history.getLastRow() === 0) history.appendRow(HISTORY_HEADERS);

  const settings = getOrCreate_(ss, SHEETS.settings);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }

  [items, count, history, settings].forEach(sh => {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#e8eef9');
    sh.autoResizeColumns(1, sh.getLastColumn());
  });

  toast_('Sheets ready. Next: Cycle Count > Generate today\'s count');
}

/* ------------------------------------------------- 1. build today's list */

function generateDailyCount() {
  const ss = SpreadsheetApp.getActive();
  const cfg = readSettings_();
  const items = readItems_();
  if (!items.length) throw new Error('The Items sheet is empty. Run "Set up sheets with sample data" first.');

  const highRisk = items.filter(i => String(i.risk).toLowerCase() === 'high');
  const others = items
    .filter(i => String(i.risk).toLowerCase() !== 'high')
    // never counted first, then the oldest count first; Medium before Low on ties
    .sort((a, b) => lastCountedTime_(a) - lastCountedTime_(b) || riskRank_(a) - riskRank_(b));

  const slots = Math.max(0, cfg.itemsPerDay - highRisk.length);
  const list = highRisk.concat(others.slice(0, slots));

  const sheet = ss.getSheetByName(SHEETS.count);
  sheet.clearContents();
  sheet.appendRow(COUNT_HEADERS);

  const rows = list.map((i, n) => {
    const r = n + 2;
    return [
      i.sku, i.name, i.location, i.systemQty, '',
      `=IF(E${r}="","",E${r}-D${r})`,
      `=IF(E${r}="","",F${r}*${Number(i.unitCost) || 0})`,
      `=IF(E${r}="","Pending",IF(OR(ABS(F${r})>=${cfg.alertUnits},ABS(G${r})>=${cfg.alertValue}),"CHECK","OK"))`,
    ];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, COUNT_HEADERS.length).setValues(rows);

  sheet.getRange('E:E').setBackground(null);
  if (rows.length) sheet.getRange(2, 5, rows.length, 1).setBackground('#fff8d6'); // the only column staff fill
  sheet.getRange('G:G').setNumberFormat('$#,##0.00');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, COUNT_HEADERS.length);
  ss.setActiveSheet(sheet);

  toast_(`${rows.length} items to count today (${highRisk.length} high-risk).`);
}

/* ----------------------------------------------- 2. save + report the count */

function submitCount() {
  const ss = SpreadsheetApp.getActive();
  const cfg = readSettings_();
  const sheet = ss.getSheetByName(SHEETS.count);
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('There is no count to submit. Generate today\'s count first.');

  const data = sheet.getRange(2, 1, last - 1, COUNT_HEADERS.length).getValues();
  const pending = data.filter(r => r[4] === '' || r[4] === null);
  if (pending.length) {
    throw new Error(`${pending.length} item(s) still have no Counted Qty: ${pending.map(r => r[0]).join(', ')}`);
  }

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'unknown';
  const costs = new Map(readItems_().map(i => [String(i.sku), Number(i.unitCost) || 0]));
  const results = data.map(r => {
    const unitCost = costs.get(String(r[0])) || 0;
    const diff = Number(r[4]) - Number(r[3]);
    const diffValue = diff * unitCost;
    const flagged = Math.abs(diff) >= cfg.alertUnits || Math.abs(diffValue) >= cfg.alertValue;
    return {
      sku: r[0], name: r[1], location: r[2], system: Number(r[3]), counted: Number(r[4]),
      diff, diffValue, status: flagged ? 'CHECK' : 'OK',
    };
  });

  // History
  const history = ss.getSheetByName(SHEETS.history);
  const histRows = results.map(x => [now, x.sku, x.name, x.location, x.system, x.counted, x.diff, x.diffValue, x.status, user]);
  history.getRange(history.getLastRow() + 1, 1, histRows.length, HISTORY_HEADERS.length).setValues(histRows);
  history.getRange('H:H').setNumberFormat('$#,##0.00');

  // Items: last counted date (and system qty if enabled)
  const itemsSheet = ss.getSheetByName(SHEETS.items);
  const itemRows = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, ITEM_HEADERS.length).getValues();
  const bySku = new Map(results.map(x => [String(x.sku), x]));
  itemRows.forEach(row => {
    const x = bySku.get(String(row[0]));
    if (!x) return;
    row[6] = now;
    if (cfg.updateSystemQty) row[3] = x.counted;
  });
  itemsSheet.getRange(2, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);

  sendReport_(results, cfg, now);

  sheet.clearContents();
  sheet.appendRow(COUNT_HEADERS);
  toast_(`Saved ${results.length} items. ${results.filter(x => x.status === 'CHECK').length} need a check. Report sent.`);
}

/* ------------------------------------------------------------- report */

function sendReport_(results, cfg, date) {
  const to = cfg.reportEmail || Session.getEffectiveUser().getEmail();
  if (!to) return;

  const flagged = results.filter(x => x.status === 'CHECK');
  const exact = results.filter(x => x.diff === 0).length;
  const accuracy = results.length ? Math.round((exact / results.length) * 1000) / 10 : 0;
  const netValue = results.reduce((s, x) => s + x.diffValue, 0);
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const day = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  const money = n => (n < 0 ? '-$' : '$') + Math.abs(n).toFixed(2);

  const rows = flagged.map(x =>
    `<tr><td>${esc_(x.sku)}</td><td>${esc_(x.name)}</td><td>${esc_(x.location)}</td>` +
    `<td align="right">${x.system}</td><td align="right">${x.counted}</td>` +
    `<td align="right"><b>${x.diff > 0 ? '+' : ''}${x.diff}</b></td><td align="right">${money(x.diffValue)}</td></tr>`
  ).join('');

  const html =
    `<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">` +
    `<h2 style="margin:0 0 8px">Cycle count ${day}</h2>` +
    `<p>${results.length} items counted · <b>${accuracy}%</b> exact · net difference <b>${money(netValue)}</b></p>` +
    (flagged.length
      ? `<p><b>${flagged.length} item(s) to check today:</b></p>` +
        `<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ccc">` +
        `<tr style="background:#e8eef9"><th>SKU</th><th>Item</th><th>Location</th><th>System</th><th>Counted</th><th>Diff</th><th>Diff $</th></tr>` +
        rows + `</table>` +
        `<p style="color:#555">Tip: recount these items today. Most differences are counting or receiving errors; ` +
        `the ones that stay after a recount are real losses.</p>`
      : `<p>All items match. 👍</p>`) +
    `</div>`;

  MailApp.sendEmail({
    to,
    subject: `Cycle count ${day}: ${flagged.length ? flagged.length + ' to check' : 'all OK'} (${accuracy}% exact)`,
    htmlBody: html,
  });
}

/* ------------------------------------------------------------ trigger */

function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'generateDailyCount')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('generateDailyCount').timeBased().everyDays(1).atHour(7).create();
  toast_("Done. Today's count list will be ready every day at 7 AM.");
}

/* ------------------------------------------------------------ helpers */

function readItems_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.items);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ITEM_HEADERS.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ sku: r[0], name: r[1], location: r[2], systemQty: r[3], unitCost: r[4], risk: r[5], lastCounted: r[6] }));
}

function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  const map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, v]) => { map[String(k).trim()] = v; });
  }
  return {
    itemsPerDay: Number(map['Items per day']) || 15,
    alertUnits: Number(map['Alert difference (units)']) || 2,
    alertValue: Number(map['Alert difference ($)']) || 25,
    reportEmail: String(map['Report email'] || '').trim(),
    updateSystemQty: String(map['Update system qty after count'] || '').toUpperCase() === 'YES',
  };
}

function lastCountedTime_(item) {
  return item.lastCounted instanceof Date ? item.lastCounted.getTime() : 0;
}

function riskRank_(item) {
  return { high: 0, medium: 1, low: 2 }[String(item.risk).toLowerCase()] ?? 3;
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'Cycle Count', 6);
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/* --------------------------------------------------------- sample data */
// Fictional store data, only used by "Set up sheets with sample data".
const SAMPLE_ITEMS = [
  ['SH-1001', 'Shapewear Bodysuit - M', 'Front rack A', 24, 38.5, 'High'],
  ['SH-1002', 'Shapewear Bodysuit - L', 'Front rack A', 18, 38.5, 'High'],
  ['SH-1003', 'Waist Trainer - S', 'Wall B', 30, 22.0, 'High'],
  ['SH-1004', 'Waist Trainer - M', 'Wall B', 26, 22.0, 'Medium'],
  ['SH-1005', 'Post-Surgery Garment - M', 'Back room', 12, 64.0, 'Medium'],
  ['SH-1006', 'Post-Surgery Garment - L', 'Back room', 9, 64.0, 'Medium'],
  ['AC-2001', 'Compression Board', 'Counter', 40, 9.5, 'Low'],
  ['AC-2002', 'Lipo Foam Pads (pair)', 'Counter', 55, 6.0, 'Low'],
  ['AC-2003', 'Silicone Nipple Covers', 'Checkout', 80, 4.25, 'High'],
  ['AC-2004', 'Firming Cream 8 oz', 'Checkout', 35, 14.0, 'High'],
  ['AC-2005', 'Detox Tea (30 bags)', 'Checkout', 42, 7.5, 'Medium'],
  ['BR-3001', 'Seamless Bra - 34B', 'Wall C', 15, 18.0, 'Low'],
  ['BR-3002', 'Seamless Bra - 36C', 'Wall C', 17, 18.0, 'Low'],
  ['BR-3003', 'Strapless Bra - 34C', 'Wall C', 11, 21.0, 'Medium'],
  ['LG-4001', 'Butt-Lift Leggings - S', 'Front rack B', 20, 26.0, 'Medium'],
  ['LG-4002', 'Butt-Lift Leggings - M', 'Front rack B', 22, 26.0, 'Medium'],
  ['LG-4003', 'Butt-Lift Leggings - L', 'Front rack B', 14, 26.0, 'Low'],
  ['VS-5001', 'Thermal Vest - M', 'Wall D', 10, 29.0, 'Low'],
  ['VS-5002', 'Thermal Vest - L', 'Wall D', 8, 29.0, 'Low'],
  ['VS-5003', 'Posture Corrector', 'Wall D', 16, 19.0, 'Medium'],
];
