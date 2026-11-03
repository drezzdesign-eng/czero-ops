/**
 * C ZERO Ops Monitor — Google Sheets backend (Apps Script web app)
 *
 * Paste this whole file into Extensions → Apps Script of a new Google Sheet,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *
 * Tabs it keeps up to date (created automatically):
 *   _data     – full dashboard data as JSON (do not edit by hand)
 *   Teams     – one row per team, current status (rewritten on every save)
 *   History   – one row per checkpoint event (appended, never cleared)
 *   Staff     – staff list
 *   Vehicles  – saved vehicles (plate, model, GPS device ID)
 */

const APP_ID = "czero-ops";
const CHUNK = 45000; // characters per cell (Sheets cell limit is 50,000)

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "read";
  if (action === "ping") return json_({ ok: true, app: APP_ID, at: Date.now() });
  const data = readData_();
  return json_({ ok: true, record: data.record, updatedAt: data.updatedAt });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad json" }); }
  if (!body || body.app !== APP_ID) return json_({ ok: false, error: "unknown app" });
  if (body.action !== "save" || !body.data) return json_({ ok: false, error: "unknown action" });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: "busy, try again" });
  try {
    const record = {
      teams: Array.isArray(body.data.teams) ? body.data.teams : [],
      staff: Array.isArray(body.data.staff) ? body.data.staff : [],
      vehicles: Array.isArray(body.data.vehicles) ? body.data.vehicles : [],
    };
    writeData_(record);
    writeTeams_(record.teams);
    appendHistory_(record.teams);
    writeStaff_(record.staff);
    writeVehicles_(record.vehicles);
    return json_({ ok: true, savedAt: Date.now() });
  } finally {
    lock.releaseLock();
  }
}

// ---------- storage ----------
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) { sh.appendRow(headers); sh.setFrozenRows(1); sh.getRange(1, 1, 1, headers.length).setFontWeight("bold"); }
  }
  return sh;
}

function readData_() {
  const sh = sheet_("_data");
  const last = sh.getLastRow();
  if (last < 2) return { record: { teams: [], staff: [], vehicles: [] }, updatedAt: null };
  const updatedAt = sh.getRange(1, 1).getValue();
  const text = sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]).join("");
  try { return { record: JSON.parse(text), updatedAt: updatedAt }; }
  catch (err) { return { record: { teams: [], staff: [], vehicles: [] }, updatedAt: updatedAt }; }
}

function writeData_(record) {
  const sh = sheet_("_data");
  const text = JSON.stringify(record);
  const rows = [];
  for (let i = 0; i < text.length; i += CHUNK) rows.push([text.slice(i, i + CHUNK)]);
  sh.clearContents();
  sh.getRange(1, 1).setValue(new Date().toISOString());
  if (rows.length) sh.getRange(2, 1, rows.length, 1).setNumberFormat("@").setValues(rows);
}

// ---------- readable tabs ----------
const TZ = "Asia/Kuala_Lumpur";
function fmt_(ms) { return ms ? Utilities.formatDate(new Date(ms), TZ, "yyyy-MM-dd HH:mm:ss") : ""; }

function writeTeams_(teams) {
  const headers = ["Team", "Status", "Tapak", "Job", "Ketua", "Pembantu", "No. Plat", "Model", "GPS Device ID", "Lat", "Lng", "Status terakhir"];
  const sh = sheet_("Teams", headers);
  const rows = teams.map(t => {
    const h = (t.history || [])[(t.history || []).length - 1];
    return [t.name || "", t.stage || "", t.site || "", t.job || "", t.leader || "", t.assistant || "",
            t.vehiclePlate || "", t.vehicleModel || "", t.gpsDeviceId || "", t.lat || "", t.lng || "", h ? fmt_(h.at) : ""];
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function appendHistory_(teams) {
  const headers = ["Masa", "Team", "Status", "Tapak", "Job", "Ketua", "Pembantu", "Auto GPS", "Lat", "Lng", "key"];
  const sh = sheet_("History", headers);
  const last = sh.getLastRow();
  const seen = {};
  if (last > 1) {
    const from = Math.max(2, last - 3000);
    sh.getRange(from, 11, last - from + 1, 1).getValues().forEach(r => { seen[r[0]] = true; });
  }
  const rows = [];
  teams.forEach(t => (t.history || []).forEach(h => {
    const key = (t.id || t.name) + "|" + h.at + "|" + h.stage;
    if (seen[key]) return;
    seen[key] = true;
    rows.push([fmt_(h.at), t.name || "", h.stage || "", h.site || "", h.job || "", h.leader || "", h.assistant || "",
               h.auto ? "ya" : "", h.lat || "", h.lng || "", key]);
  }));
  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function writeStaff_(staff) {
  const headers = ["Nama", "Jawatan"];
  const sh = sheet_("Staff", headers);
  const rows = staff.map(s => [s.name || "", s.role || ""]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function writeVehicles_(vehicles) {
  const headers = ["No. Plat", "Model", "GPS Device ID", "Team terakhir"];
  const sh = sheet_("Vehicles", headers);
  const rows = vehicles.map(v => [v.plate || "", v.model || "", v.gpsDeviceId || "", v.lastTeam || ""]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
