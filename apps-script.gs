// Paste into Google Sheet → Extensions → Apps Script. Deploy → New deployment → Web app
// Execute as: Me · Who has access: Anyone. Copy the /exec URL into CONFIG.sheetUrl in index.html.
const COLS = ["submitted_at","invited_as","name","whatsapp","side","attendance","missing_events","partner_joining",
              "partner_name","guests","city","arrival","departure","notes"];

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVPs")
            || SpreadsheetApp.getActiveSpreadsheet().insertSheet("RSVPs");
    if (sh.getLastRow() === 0) { sh.appendRow(COLS); sh.setFrozenRows(1); }
    const d = JSON.parse(e.postData.contents);
    d.submitted_at = new Date();
    // keep WhatsApp as text so leading + / 0 survive
    sh.appendRow(COLS.map(c => c === "whatsapp" ? "'" + (d[c] || "") : (d[c] ?? "")));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

function doGet() { return ContentService.createTextOutput("RSVP endpoint is live."); }
