// Paste into Google Sheet → Extensions → Apps Script. Deploy → New deployment → Web app
// Execute as: Me · Who has access: Anyone. Copy the /exec URL into CONFIG.sheetUrl in index.html.
// After changing this file: Deploy → Manage deployments → Edit → Version: New version (the URL stays the same).
const COLS = ["submitted_at","invited_as","name","whatsapp","side","attendance","missing_events","partner_joining",
              "partner_name","guests","city","arrival","departure","notes","how_known"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("RSVPs") || ss.insertSheet("RSVPs");
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_();
    // new columns are added at the end, so the header row is rewritten when it is shorter than COLS
    if (sh.getLastRow() === 0 || sh.getLastColumn() < COLS.length) {
      sh.getRange(1, 1, 1, COLS.length).setValues([COLS]); sh.setFrozenRows(1);
    }
    const d = JSON.parse(e.postData.contents);
    d.submitted_at = new Date();
    // keep WhatsApp as text so leading + / 0 survive
    sh.appendRow(COLS.map(c => c === "whatsapp" ? "'" + (d[c] || "") : (d[c] ?? "")));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

// ?list=1 → first name + last initial of everyone coming (latest reply per WhatsApp number wins)
function doGet(e) {
  if (!(e && e.parameter && e.parameter.list)) return ContentService.createTextOutput("RSVP endpoint is live.");
  const rows = sheet_().getDataRange().getValues();
  const h = rows.shift() || [], col = k => h.indexOf(k);
  const latest = {};
  rows.forEach(r => { const k = String(r[col("whatsapp")]).replace(/\D/g, "") || r[col("name")]; latest[k] = r; });
  const names = Object.values(latest)
    .filter(r => r[col("attendance")] && r[col("attendance")] !== "Can't make it")
    .map(r => { const [f, ...rest] = String(r[col("name")]).trim().split(/\s+/); const l = rest.pop(); return l ? `${f} ${l[0]}.` : f; })
    .filter(n => n && !/^test\b/i.test(n));
  return ContentService.createTextOutput(JSON.stringify({ names })).setMimeType(ContentService.MimeType.JSON);
}
