const SPREADSHEET_ID = "1Td2IZQiZiRHwptywxUAeES8cljTH-6i4eFoNrFmmdLY";
const SHEET_NAME = "Guests";

const PASS_TYPES = [
  "Regular",
  "Early Bird",
  "Couple",
  "Surge Pass",
  "Guest List",
  "Organizers"
];

function doGet(e) {
  const action = e.parameter.action || "";

  if (action === "generate") return generatePass_(e);
  if (action === "check")    return checkPass_(e);
  if (action === "approve")  return approvePass_(e);

  return json_({ ok: false, message: "Invalid action." });
}

function generatePass_(e) {
  const name = (e.parameter.name || "").trim();
  const type = (e.parameter.type || "").trim();

  if (!name) return json_({ ok: false, message: "Guest name is required." });
  if (!PASS_TYPES.includes(type)) return json_({ ok: false, message: "Invalid pass type." });

  const sheet = getSheet_();

  let passId;
  do {
    passId = "VYRA-" + Math.random().toString(36).substring(2, 8).toUpperCase();
  } while (findRow_(sheet, passId) !== -1);

  sheet.appendRow([passId, name, type, passId, "Not Used", new Date()]);

  return json_({
    ok: true,
    pass_id: passId,
    guest_name: name,
    pass_type: type,
    qr_data: passId
  });
}

function checkPass_(e) {
  const id = normalizeId_(e.parameter.id || "");
  if (!id) return json_({ ok: false, message: "Pass ID is required." });

  const sheet = getSheet_();
  const row = findRow_(sheet, id);

  if (row === -1) return json_({ ok: true, valid: false, message: "Invalid pass." });

  const v = sheet.getRange(row, 1, 1, 6).getValues()[0];

  return json_({
    ok: true,
    valid: true,
    pass_id: v[0],
    guest_name: v[1],
    pass_type: v[2],
    status: v[4]
  });
}

function approvePass_(e) {
  const id = normalizeId_(e.parameter.id || "");
  if (!id) return json_({ ok: false, message: "Pass ID is required." });

  const sheet = getSheet_();
  const row = findRow_(sheet, id);

  if (row === -1) return json_({ ok: false, message: "Pass not found." });

  sheet.getRange(row, 5).setValue("Used");

  return json_({ ok: true, message: "Pass approved.", pass_id: id });
}

function normalizeId_(id) {
  return String(id).trim().toUpperCase();
}

function findRow_(sheet, passId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const target = normalizeId_(passId);

  for (let i = 0; i < ids.length; i++) {
    if (normalizeId_(ids[i][0]) === target) return i + 2;
  }
  return -1;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
