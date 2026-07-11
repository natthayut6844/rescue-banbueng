/**
 * ระบบบันทึกการเข้าเวรกู้ชีพกู้ภัย — หน่วยกู้ชีพกู้ภัยตำบลบ้านบึง
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีติดตั้ง (สรุป — ดูรายละเอียดใน คู่มือติดตั้ง.md):
 * 1) สร้าง Google Sheet ใหม่ > Extensions > Apps Script > วางโค้ดนี้ทั้งหมด
 * 2) รันฟังก์ชัน setup() หนึ่งครั้ง (สร้างชีทและแอดมินเริ่มต้น: id=admin, pin=1234)
 * 3) Deploy > New deployment > Web app > Execute as: Me / Access: Anyone
 * 4) คัดลอก URL ที่ได้ ไปใส่ในตัวแปร API_URL ของไฟล์ index.html
 */

const TZ = 'Asia/Bangkok';
const SHIFT_HOURS = 12;            // ความยาวเวร (ชั่วโมง)
const EARLY_GRACE_MIN = 15;        // ออกก่อนเวลาเกินกี่นาทีถึงถือว่า "ออกก่อนเวลา"
const LATE_GRACE_MIN = 15;         // เข้าช้ากว่ากำหนดเกินกี่นาทีถึงถือว่า "เข้าสาย"
const SELFIE_FOLDER = 'RescueApp_Selfies';

// ---------- Web entry ----------
function doGet(e) {
  return json({ ok: true, service: 'BanBueng Rescue Duty API', time: new Date().toISOString() });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: 'คำขอไม่ถูกต้อง' }); }
  try {
    return json(route(req));
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Router ----------
function route(req) {
  const a = req.action;
  if (a === 'login') return login(req);

  // ทุก action ต่อจากนี้ต้องยืนยันตัวตน
  const user = auth(req.id, req.pin);
  if (!user) return { ok: false, error: 'รหัสหรือ PIN ไม่ถูกต้อง' };
  const isAdmin = user.role === 'admin';

  switch (a) {
    case 'checkin':        return checkin(user, req);
    case 'checkout':       return checkout(user, req);
    case 'ping':           return ping(user, req);
    case 'board':          return board();
    case 'stats':          return { ok: true, data: statsMonth(req.month) };
    case 'payroll':        return payroll(req.month);
    case 'myinfo':         return myinfo(user, req.month);
    case 'history':        return history(req.userId || user.id, req.date, user);
    case 'outbase':        return outbase(user, req);
    case 'confirmDispatch':return confirmDispatch(user, req);
    case 'endDispatch':    return endDispatch(user, req);
    case 'events':         return listEvents(req.month);
    case 'getSettings':    return { ok: true, data: getSettings() };
    case 'saveSettings':   return isAdmin ? saveSettings(req.data) : deny();
    case 'listUsers':      return isAdmin ? { ok: true, data: listUsers() } : deny();
    case 'saveUser':       return isAdmin ? saveUser(req.data) : deny();
    case 'deleteUser':     return isAdmin ? deleteUser(req.userId) : deny();
    default: return { ok: false, error: 'ไม่รู้จักคำสั่ง: ' + a };
  }
}
function deny() { return { ok: false, error: 'สิทธิ์แอดมินเท่านั้น' }; }

// ---------- Sheets ----------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name, headers) {
  let sh = ss().getSheetByName(name);
  if (!sh) { sh = ss().insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function shUsers()  { return sheet('Users',    ['id','pin','name','role','phone','createdAt']); }
function shChecks() { return sheet('Checks',   ['timestamp','userId','name','action','shiftDate','shiftType','lat','lng','selfieUrl','reason']); }
function shLocs()   { return sheet('Locations',['timestamp','userId','lat','lng','distanceM']); }
function shEvents() { return sheet('Events',   ['timestamp','eventId','userId','name','type','detail','shiftDate','shiftType','status','confirmedBy']); }
function shSet()    { return sheet('Settings', ['key','value']); }

function rows(sh) {
  const v = sh.getDataRange().getValues();
  return v.slice(1);
}

/** รันครั้งเดียวหลังวางโค้ด: สร้างชีท + แอดมินเริ่มต้น + ค่าตั้งต้น */
function setup() {
  shUsers(); shChecks(); shLocs(); shEvents(); shSet();
  if (listUsers().length === 0) {
    shUsers().appendRow(['admin', '1234', 'ผู้ดูแลระบบ', 'admin', '', new Date()]);
  }
  const s = getSettings();
  if (!s.radiusM) setSetting('radiusM', 300);
  if (s.prorate === undefined) setSetting('prorate', '0');
}

// ---------- Auth / Users ----------
function auth(id, pin) {
  if (!id || pin === undefined) return null;
  const u = listUsers().find(x => String(x.id) === String(id) && String(x.pin) === String(pin));
  return u || null;
}
function login(req) {
  const u = auth(req.id, req.pin);
  if (!u) return { ok: false, error: 'รหัสหรือ PIN ไม่ถูกต้อง' };
  return { ok: true, user: { id: u.id, name: u.name, role: u.role, phone: u.phone } };
}
function listUsers() {
  return rows(shUsers()).filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), pin: String(r[1]), name: r[2], role: r[3] || 'staff', phone: r[4] || ''
  }));
}
function saveUser(d) {
  if (!d || !d.id || !d.name) return { ok: false, error: 'ข้อมูลไม่ครบ' };
  const sh = shUsers();
  const data = rows(sh);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(d.id)) {
      sh.getRange(i + 2, 1, 1, 5).setValues([[d.id, d.pin || data[i][1], d.name, d.role || 'staff', d.phone || '']]);
      return { ok: true };
    }
  }
  sh.appendRow([d.id, d.pin || '0000', d.name, d.role || 'staff', d.phone || '', new Date()]);
  return { ok: true };
}
function deleteUser(userId) {
  const sh = shUsers();
  const data = rows(sh);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) { sh.deleteRow(i + 2); return { ok: true }; }
  }
  return { ok: false, error: 'ไม่พบผู้ใช้' };
}

// ---------- Settings ----------
function getSettings() {
  const o = {};
  rows(shSet()).forEach(r => { if (r[0] !== '') o[r[0]] = r[1]; });
  return o;
}
function setSetting(key, value) {
  const sh = shSet();
  const data = rows(sh);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}
function saveSettings(data) {
  if (!data) return { ok: false, error: 'ไม่มีข้อมูล' };
  Object.keys(data).forEach(k => setSetting(k, data[k]));
  return { ok: true, data: getSettings() };
}

// ---------- เวลา / กะเวร ----------
function fmt(d, p) { return Utilities.formatDate(new Date(d), TZ, p); }
function dstr(d) { return fmt(d, 'yyyy-MM-dd'); }
function tstr(d) { return fmt(d, 'HH:mm'); }

/** คืนกะเวรของเวลาที่ให้มา: กะเช้า 06:00–18:00 / กะบ่าย(คืน) 18:00–06:00 ของวันถัดไป */
function shiftOf(d) {
  d = new Date(d);
  const h = Number(fmt(d, 'H'));
  let shiftDate, type;
  if (h >= 6 && h < 18) { shiftDate = dstr(d); type = 'day'; }
  else if (h >= 18)     { shiftDate = dstr(d); type = 'night'; }
  else { // 00:00–05:59 = กะบ่ายของเมื่อวาน
    const prev = new Date(d.getTime() - 24 * 3600 * 1000);
    shiftDate = dstr(prev); type = 'night';
  }
  const start = new Date(shiftDate + 'T' + (type === 'day' ? '06:00:00' : '18:00:00') + '+07:00');
  const end = new Date(start.getTime() + SHIFT_HOURS * 3600 * 1000);
  return { shiftDate: shiftDate, type: type, start: start, end: end };
}
function shiftLabel(t) { return t === 'day' ? 'กะเช้า (06:00–18:00)' : 'กะบ่าย (18:00–06:00)'; }

// ---------- เช็คอิน / เช็คเอาท์ ----------
function saveSelfie(b64, userId) {
  try {
    let folder;
    const it = DriveApp.getFoldersByName(SELFIE_FOLDER);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(SELFIE_FOLDER);
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', userId + '_' + Date.now() + '.jpg');
    const f = folder.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w240';
  } catch (e) { return ''; }
}

function userShiftChecks(userId, shiftDate, type) {
  return rows(shChecks()).filter(r =>
    String(r[1]) === String(userId) && String(r[4]) === String(shiftDate) && r[5] === type
  ).map(rowToCheck).sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
}
function rowToCheck(r) {
  return { timestamp: r[0] instanceof Date ? r[0].toISOString() : r[0], userId: String(r[1]), name: r[2],
           action: r[3], shiftDate: r[4] instanceof Date ? dstr(r[4]) : String(r[4]), shiftType: r[5],
           lat: r[6], lng: r[7], selfieUrl: r[8], reason: r[9] || '' };
}

function checkin(user, req) {
  const now = new Date();
  const s = shiftOf(now);
  const prev = userShiftChecks(user.id, s.shiftDate, s.type);
  const last = prev[prev.length - 1];
  if (last && last.action === 'checkin') return { ok: false, error: 'คุณเช็คอินค้างอยู่แล้ว' };

  const selfieUrl = req.selfie ? saveSelfie(req.selfie, user.id) : '';
  shChecks().appendRow([now, user.id, user.name, 'checkin', s.shiftDate, s.type, req.lat || '', req.lng || '', selfieUrl, req.reason || '']);
  if (req.lat) logLoc(user.id, req.lat, req.lng);

  const count = prev.filter(c => c.action === 'checkin').length + 1;
  const late = now.getTime() > s.start.getTime() + LATE_GRACE_MIN * 60000;
  return { ok: true, shift: { date: s.shiftDate, type: s.type, label: shiftLabel(s.type), start: tstr(s.start), end: tstr(s.end) },
           time: tstr(now), checkinCount: count, late: late, selfieUrl: selfieUrl };
}

function checkout(user, req) {
  const now = new Date();
  // หากะที่เช็คอินค้างอยู่ (กะปัจจุบัน หรือกะก่อนหน้า เผื่อลืมเช็คเอาท์)
  const cands = [shiftOf(now), shiftOf(new Date(now.getTime() - 12 * 3600 * 1000))];
  let s = null, checks = null;
  for (const c of cands) {
    const list = userShiftChecks(user.id, c.shiftDate, c.type);
    const last = list[list.length - 1];
    if (last && last.action === 'checkin') { s = c; checks = list; break; }
  }
  if (!s) return { ok: false, error: 'คุณยังไม่ได้เช็คอิน' };

  const early = now.getTime() < s.end.getTime() - EARLY_GRACE_MIN * 60000;
  if (early && !req.reason) {
    // ให้ฝั่งแอปเปิดหน้าต่างแจ้งสาเหตุ + ยืนยันเวลา
    return { ok: false, needReason: true,
             inTime: tstr(checks[0].timestamp), outTime: tstr(now),
             shiftEnd: tstr(s.end), error: 'ออกเวรก่อนเวลา กรุณาแจ้งสาเหตุ' };
  }
  shChecks().appendRow([now, user.id, user.name, 'checkout', s.shiftDate, s.type, req.lat || '', req.lng || '', '', req.reason || '']);
  const hrs = shiftHours(userShiftChecks(user.id, s.shiftDate, s.type), s);
  return { ok: true, inTime: tstr(checks[0].timestamp), outTime: tstr(now), early: early,
           hours: Math.round(hrs * 100) / 100, shift: { date: s.shiftDate, type: s.type } };
}

/** คำนวณชั่วโมงที่อยู่เวรจริงจากคู่ checkin/checkout ภายในกะ */
function shiftHours(checks, s) {
  let total = 0, inAt = null;
  const now = new Date();
  checks.forEach(c => {
    const t = new Date(c.timestamp);
    if (c.action === 'checkin' && !inAt) inAt = t;
    if (c.action === 'checkout' && inAt) { total += (t - inAt); inAt = null; }
  });
  if (inAt) { // ยังค้างเช็คอิน
    const end = now < s.end ? now : s.end;
    total += Math.max(0, end - inAt);
  }
  return total / 3600000;
}

// ---------- พิกัด ----------
function logLoc(userId, lat, lng) {
  const st = getSettings();
  let dist = '';
  if (st.baseLat && st.baseLng) dist = Math.round(haversine(Number(lat), Number(lng), Number(st.baseLat), Number(st.baseLng)));
  shLocs().appendRow([new Date(), userId, lat, lng, dist]);
  return dist;
}
function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
function ping(user, req) {
  if (!req.lat) return { ok: false, error: 'ไม่มีพิกัด' };
  const dist = logLoc(user.id, req.lat, req.lng);
  const st = getSettings();
  const radius = Number(st.radiusM || 300);
  const outside = dist !== '' && dist > radius;
  return { ok: true, distanceM: dist === '' ? null : dist, radiusM: radius, outside: outside,
           dispatch: activeDispatch() ? true : false };
}
function history(userId, date, requester) {
  // เจ้าหน้าที่ดูของตัวเองได้ / แอดมินดูได้ทุกคน
  if (requester.role !== 'admin' && String(userId) !== String(requester.id)) return deny();
  const day = date || dstr(new Date());
  const pts = rows(shLocs()).filter(r => String(r[1]) === String(userId) && dstr(r[0]) === day)
    .map(r => ({ time: tstr(r[0]), lat: r[2], lng: r[3], distanceM: r[4] }));
  return { ok: true, data: pts, date: day };
}

// ---------- เหตุการณ์ออกนอกศูนย์ / ออกเหตุ ----------
function outbase(user, req) {
  const now = new Date();
  const s = shiftOf(now);
  const eventId = 'E' + now.getTime();
  const type = req.type || 'other'; // dispatch = ออกเหตุ
  const status = type === 'dispatch' ? 'pending' : 'reported';
  shEvents().appendRow([now, eventId, user.id, user.name, type, req.detail || '', s.shiftDate, s.type, status, '']);
  return { ok: true, eventId: eventId, status: status };
}
function activeDispatch() {
  const s = shiftOf(new Date());
  const ev = rows(shEvents()).filter(r =>
    r[4] === 'dispatch' && String(r[6]) === s.shiftDate && r[7] === s.type &&
    (r[8] === 'pending' || r[8] === 'confirmed'));
  return ev.length ? { eventId: ev[ev.length - 1][1], status: ev[ev.length - 1][8],
                       by: ev[ev.length - 1][3], detail: ev[ev.length - 1][5],
                       confirmedBy: ev[ev.length - 1][9], time: tstr(ev[ev.length - 1][0]) } : null;
}
/** คนใดคนหนึ่งในเวรกดยืนยัน = ทุกคนในเวรถือว่าออกปฏิบัติงานอัตโนมัติ */
function confirmDispatch(user, req) {
  const sh = shEvents();
  const data = rows(sh);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(req.eventId)) {
      sh.getRange(i + 2, 9).setValue('confirmed');
      sh.getRange(i + 2, 10).setValue(user.name);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบเหตุการณ์' };
}
function endDispatch(user, req) {
  const sh = shEvents();
  const data = rows(sh);
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][4] === 'dispatch' && (data[i][8] === 'confirmed' || data[i][8] === 'pending')) {
      sh.getRange(i + 2, 9).setValue('closed');
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่มีเหตุที่เปิดอยู่' };
}
function listEvents(month) {
  const m = month || fmt(new Date(), 'yyyy-MM');
  const ev = rows(shEvents()).filter(r => fmt(r[0], 'yyyy-MM') === m)
    .map(r => ({ time: fmt(r[0], 'dd/MM HH:mm'), eventId: r[1], name: r[3], type: r[4],
                 detail: r[5], shiftDate: r[6] instanceof Date ? dstr(r[6]) : r[6],
                 shiftType: r[7], status: r[8], confirmedBy: r[9] }))
    .reverse();
  return { ok: true, data: ev };
}

// ---------- แดชบอร์ด "ผู้ปฏิบัติงานวันนี้" ----------
function board() {
  const now = new Date();
  const s = shiftOf(now);
  const st = getSettings();
  const checksAll = rows(shChecks()).filter(r => String(r[4] instanceof Date ? dstr(r[4]) : r[4]) === s.shiftDate && r[5] === s.type).map(rowToCheck);
  const byUser = {};
  checksAll.forEach(c => { (byUser[c.userId] = byUser[c.userId] || []).push(c); });

  const dispatch = activeDispatch();
  const locsToday = rows(shLocs()).filter(r => new Date(r[0]) >= s.start);

  const people = Object.keys(byUser).map(uid => {
    const list = byUser[uid].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
    const last = list[list.length - 1];
    const ins = list.filter(c => c.action === 'checkin');
    const selfie = (ins[ins.length - 1] || {}).selfieUrl || (ins[0] || {}).selfieUrl || '';
    const myLocs = locsToday.filter(r => String(r[1]) === String(uid));
    const lastLoc = myLocs[myLocs.length - 1];
    const onDuty = last.action === 'checkin';
    let status = onDuty ? 'on' : 'off';
    if (onDuty && dispatch && dispatch.status === 'confirmed') status = 'dispatch';
    else if (onDuty && lastLoc && lastLoc[4] !== '' && Number(lastLoc[4]) > Number(st.radiusM || 300)) status = 'outside';
    return {
      userId: uid, name: last.name, selfieUrl: selfie, status: status,
      firstIn: tstr(list[0].timestamp), lastAction: tstr(last.timestamp),
      checkinCount: ins.length,
      lastLoc: lastLoc ? { time: tstr(lastLoc[0]), lat: lastLoc[2], lng: lastLoc[3], distanceM: lastLoc[4] } : null,
      hours: Math.round(shiftHours(list, s) * 100) / 100,
      reason: list.filter(c => c.reason).map(c => c.reason).join(' / ')
    };
  });
  people.sort((a, b) => (a.status === 'off') - (b.status === 'off'));
  return { ok: true,
    shift: { date: s.shiftDate, type: s.type, label: shiftLabel(s.type), start: tstr(s.start), end: tstr(s.end) },
    dispatch: dispatch, base: { lat: st.baseLat || '', lng: st.baseLng || '', radiusM: st.radiusM || 300 },
    people: people, serverTime: fmt(now, 'HH:mm:ss') };
}

// ---------- สถิติ / ค่าตอบแทน ----------
/** สถิติรายเดือน: ต่อคน = จำนวนเวร, ชั่วโมงรวม, เวรไม่ครบ, เช็คอินซ้ำ */
function statsMonth(month) {
  const m = month || fmt(new Date(), 'yyyy-MM');
  const checks = rows(shChecks()).map(rowToCheck).filter(c => String(c.shiftDate).slice(0, 7) === m);
  const key = c => c.userId + '|' + c.shiftDate + '|' + c.shiftType;
  const groups = {};
  checks.forEach(c => { (groups[key(c)] = groups[key(c)] || []).push(c); });

  const perUser = {};
  Object.keys(groups).forEach(k => {
    const list = groups[k].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
    const [uid, sd, tp] = k.split('|');
    const start = new Date(sd + 'T' + (tp === 'day' ? '06:00:00' : '18:00:00') + '+07:00');
    const s = { start: start, end: new Date(start.getTime() + SHIFT_HOURS * 3600000) };
    const hrs = shiftHours(list, s);
    const ins = list.filter(c => c.action === 'checkin').length;
    const incomplete = hrs < SHIFT_HOURS - 0.5;
    const u = perUser[uid] = perUser[uid] || { userId: uid, name: list[0].name, shifts: 0, hours: 0, incomplete: 0, reCheckins: 0, unit: 0, detail: [] };
    u.shifts += 1; u.hours += hrs;
    if (incomplete) u.incomplete += 1;
    if (ins > 1) u.reCheckins += 1;
    u.unit += Math.min(hrs / SHIFT_HOURS, 1);
    u.detail.push({ shiftDate: sd, shiftType: tp, hours: Math.round(hrs * 100) / 100,
                    checkinCount: ins, incomplete: incomplete,
                    inTime: tstr(list[0].timestamp),
                    outTime: list[list.length - 1].action === 'checkout' ? tstr(list[list.length - 1].timestamp) : '-',
                    reason: list.filter(c => c.reason).map(c => c.reason).join(' / ') });
  });
  const out = Object.values(perUser);
  out.forEach(u => { u.hours = Math.round(u.hours * 100) / 100; u.unit = Math.round(u.unit * 1000) / 1000;
                     u.detail.sort((a, b) => (a.shiftDate > b.shiftDate ? -1 : 1)); });
  out.sort((a, b) => b.shifts - a.shifts);
  return { month: m, users: out };
}

/**
 * ค่าตอบแทน: ยอดเงินสนับสนุนจาก อบต. (แอดมินตั้งค่ารายเดือน)
 * อัตราต่อเวร = ยอดเงินเดือนนั้น ÷ จำนวนเวรรวมทุกคนที่เข้าในเดือนนั้น
 * (ถ้าตั้งค่า "คิดตามสัดส่วนชั่วโมง" เวรที่เข้าไม่ครบจะได้ตามสัดส่วนชั่วโมงจริง)
 */
function payroll(month) {
  const m = month || fmt(new Date(), 'yyyy-MM');
  const st = getSettings();
  const budget = Number(st['budget_' + m] || 0);
  const prorate = String(st.prorate) === '1';
  const stats = statsMonth(m);
  const totalUnits = stats.users.reduce((s, u) => s + (prorate ? u.unit : u.shifts), 0);
  const rate = totalUnits > 0 ? budget / totalUnits : 0;
  const list = stats.users.map(u => ({
    userId: u.userId, name: u.name, shifts: u.shifts, hours: u.hours,
    unit: prorate ? u.unit : u.shifts,
    pay: Math.round((prorate ? u.unit : u.shifts) * rate * 100) / 100
  }));
  return { ok: true, data: { month: m, budget: budget, prorate: prorate, totalUnits: Math.round(totalUnits * 1000) / 1000,
           ratePerShift: Math.round(rate * 100) / 100, users: list } };
}

function myinfo(user, month) {
  const m = month || fmt(new Date(), 'yyyy-MM');
  const stats = statsMonth(m);
  const me = stats.users.find(u => String(u.userId) === String(user.id)) ||
             { userId: user.id, name: user.name, shifts: 0, hours: 0, incomplete: 0, reCheckins: 0, unit: 0, detail: [] };
  const pr = payroll(m).data;
  const myPay = (pr.users.find(u => String(u.userId) === String(user.id)) || {}).pay || 0;
  // สถานะปัจจุบัน
  const now = new Date();
  const cands = [shiftOf(now), shiftOf(new Date(now.getTime() - 12 * 3600000))];
  let duty = null;
  for (const c of cands) {
    const list = userShiftChecks(user.id, c.shiftDate, c.type);
    const last = list[list.length - 1];
    if (last && last.action === 'checkin') {
      duty = { shiftDate: c.shiftDate, type: c.type, label: shiftLabel(c.type),
               inTime: tstr(list[0].timestamp), end: tstr(c.end),
               checkinCount: list.filter(x => x.action === 'checkin').length };
      break;
    }
  }
  return { ok: true, data: { month: m, me: me, pay: myPay, ratePerShift: pr.ratePerShift, duty: duty } };
}
