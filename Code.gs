/**
 * VYRA PartyPass backend (Google Apps Script)
 * Routes (all GET, JSON):
 *   ?action=generate&name=..&type=..   -> generatePass()
 *   ?action=check&pass_id=PASS-XXXXXXXX
 *   ?action=approve&pass_id=PASS-XXXXXXXX
 */
const SPREADSHEET_ID = "1BYdbK0VjR-RfzMg6PyUda-H0HyE-camtus4lopLVFbQ";
const BACKGROUND_FILE_ID = "1ZiXsjkVLeZBncyDKpsvhoawx_TinVsQ9";
const SHEET_NAME = "Guests";
const PASS_TYPES = ["Regular", "Early Bird", "Couple", "Surge Pass"];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || "").toLowerCase();
  try {
    if (action === "check") return json_(checkPass_(p.pass_id));
    if (action === "approve") return json_(approvePass_(p.pass_id));
    if (action === "generate") return json_(generatePass(p.name, p.type));
    return ContentService.createTextOutput("VYRA PartyPass backend is running.");
  } catch (err) {
    return json_({ ok: false, result: "error", message: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getGuestsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Pass ID", "Guest Name", "Pass Type", "QR / Pass ID", "Status", "Created At"]);
  }
  return sh;
}

// Uppercase, trim, and collapse any repeated PASS- prefix into exactly one.
function normalizeId_(id) {
  let s = String(id || "").trim().toUpperCase();
  if (!s) return "";
  s = s.replace(/^(PASS-)+/, "");
  return "PASS-" + s;
}

// Returns the 1-based sheet row for a pass ID, or -1 (header row excluded).
function findPassRow_(sheet, passId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (normalizeId_(ids[i][0]) === passId) return i + 2;
  }
  return -1;
}

function checkPass_(rawId) {
  const passId = normalizeId_(rawId);
  if (!/^PASS-[A-Z0-9]{8}$/.test(passId)) {
    return { result: "invalid", message: "This pass is not registered.", pass_id: passId };
  }
  const sh = getGuestsSheet_();
  const row = findPassRow_(sh, passId);
  if (row < 0) return { result: "invalid", message: "This pass is not registered.", pass_id: passId };
  const v = sh.getRange(row, 1, 1, 5).getValues()[0];
  const base = { pass_id: passId, guest_name: String(v[1]), pass_type: String(v[2]) };
  if (String(v[4]).trim().toLowerCase() === "used") {
    return Object.assign({ result: "already", message: "This pass has already been checked in." }, base);
  }
  return Object.assign({ result: "valid", message: "Valid pass. Entry can be approved." }, base);
}

function approvePass_(rawId) {
  const passId = normalizeId_(rawId);
  if (!/^PASS-[A-Z0-9]{8}$/.test(passId)) {
    return { result: "invalid", message: "This pass is not registered.", pass_id: passId };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getGuestsSheet_();
    const row = findPassRow_(sh, passId);
    if (row < 0) return { result: "invalid", message: "This pass is not registered.", pass_id: passId };
    const v = sh.getRange(row, 1, 1, 5).getValues()[0];
    const base = { pass_id: passId, guest_name: String(v[1]), pass_type: String(v[2]) };
    if (String(v[4]).trim().toLowerCase() === "used") {
      return Object.assign({ result: "already", message: "This pass has already been checked in." }, base);
    }
    sh.getRange(row, 5).setValue("Used");
    SpreadsheetApp.flush();
    return Object.assign({ result: "approved", message: "Entry approved." }, base);
  } finally {
    lock.releaseLock();
  }
}

function generatePass(name, passType) {
  // Validate
  let guest = String(name || "").replace(/\s+/g, " ").trim();
  if (!guest) return { ok: false, message: "Guest name is required." };
  if (guest.length > 60) return { ok: false, message: "Guest name is too long (max 60)." };
  guest = guest.replace(/^[=+\-@]+/, ""); // prevent sheet formula injection
  if (!guest) return { ok: false, message: "Enter a valid guest name." };
  const type = PASS_TYPES.filter(function (t) { return t.toLowerCase() === String(passType || "").trim().toLowerCase(); })[0];
  if (!type) return { ok: false, message: "Invalid pass type." };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let passId;
  try {
    const sh = getGuestsSheet_();
    do {
      passId = "PASS-" + Utilities.getUuid().replace(/-/g, "").substring(0, 8).toUpperCase();
    } while (findPassRow_(sh, passId) > 0);
    sh.appendRow([passId, guest, type, passId, "Not Used", new Date()]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // QR contains ONLY the pass ID
  const qrBlob = UrlFetchApp.fetch(
    "https://api.qrserver.com/v1/create-qr-code/?size=500x500&margin=0&format=png&data=" + encodeURIComponent(passId)
  ).getBlob();
  const qr = "data:image/png;base64," + Utilities.base64Encode(qrBlob.getBytes());

  let background = "";
  try {
    const bg = DriveApp.getFileById(BACKGROUND_FILE_ID).getBlob();
    background = "data:" + bg.getContentType() + ";base64," + Utilities.base64Encode(bg.getBytes());
  } catch (err) { /* ticket still renders with default backdrop */ }

  return { ok: true, pass_id: passId, guest_name: guest, pass_type: type, qr: qr, background: background };
}
