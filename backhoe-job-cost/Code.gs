/**
 * หลังบ้านแอป "คิดต้นทุนงานดิน" — ซิงก์ข้อมูลหลายเครื่อง
 * Backend: Google Apps Script + Google Sheets (last-write-wins ต่อรายการ)
 *
 * วิธีติดตั้ง:
 * 1) สร้าง Google Sheet ใหม่ > Extensions > Apps Script > ลบโค้ดเดิม วางโค้ดนี้ทั้งหมด
 * 2) แก้ TEAM_CODE กับ PIN ข้างล่างเป็นของทีมตัวเอง แล้วรันฟังก์ชัน setup() หนึ่งครั้ง
 * 3) Deploy > New deployment > Web app > Execute as: Me / Who has access: Anyone
 * 4) คัดลอก URL ที่ได้ (ลงท้าย /exec) ไปวางในแอป: หน้า เงิน > ปุ่มตั้งค่า (ฟันเฟือง)
 *    > ช่อง API URL + รหัสทีม + PIN แล้วกด "บันทึก + ซิงก์"
 * 5) ทุกเครื่องที่ใส่ URL + รหัสทีม + PIN เดียวกัน จะเห็นข้อมูลชุดเดียวกัน
 */

const TEAM_CODE = 'ทีมเรา';   // ← แก้เป็นรหัสทีมตัวเอง
const PIN = '1234';           // ← แก้เป็น PIN ตัวเอง

const SHEET_NAME = 'Records';
const HEADERS = ['key', 'kind', 'id', 'json', 'updatedAt', 'deleted'];

function setup() {
  sheetRecords();
  Logger.log('พร้อมใช้แล้ว — Deploy เป็น Web app ได้เลย');
}

function doGet() {
  return json({ ok: true, service: 'Backhoe Job Cost Sync API', time: new Date().toISOString() });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: 'คำขอไม่ถูกต้อง' }); }
  try {
    if (String(req.teamCode) !== TEAM_CODE || String(req.pin) !== PIN)
      return json({ ok: false, error: 'รหัสทีมหรือ PIN ไม่ถูกต้อง' });
    if (req.action === 'sync') return json(sync(req));
    return json({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + req.action });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheetRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

/**
 * sync: รับ records ที่เครื่องลูกแก้มา (push) — เก็บแบบ last-write-wins ต่อรายการ
 * แล้วส่งกลับทุกรายการที่เปลี่ยนหลัง since (pull)
 */
function sync(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheetRecords();
    const data = sh.getDataRange().getValues(); // row 0 = headers
    const rowByKey = {};
    for (let i = 1; i < data.length; i++) rowByKey[data[i][0]] = i; // key -> row index

    // --- push (last-write-wins ต่อ record ตาม updatedAt) ---
    (req.records || []).forEach(function (r) {
      if (!r || !r.kind || !r.id) return;
      const key = r.kind + ':' + r.id;
      const upd = Number(r.updatedAt) || Date.now();
      const row = [key, r.kind, r.id, r.json || '', upd, r.deleted ? 1 : ''];
      if (key in rowByKey) {
        const i = rowByKey[key];
        const existing = Number(data[i][4]) || 0;
        if (upd >= existing) {
          sh.getRange(i + 1, 1, 1, HEADERS.length).setValues([row]);
          data[i] = row;
        }
      } else {
        sh.appendRow(row);
        rowByKey[key] = data.length;
        data.push(row);
      }
    });

    // --- pull ---
    const since = Number(req.since) || 0;
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const upd = Number(data[i][4]) || 0;
      if (upd > since) {
        out.push({ kind: data[i][1], id: data[i][2], json: data[i][3], updatedAt: upd, deleted: !!data[i][5] });
      }
    }
    return { ok: true, records: out, serverTime: Date.now() };
  } finally {
    lock.releaseLock();
  }
}
