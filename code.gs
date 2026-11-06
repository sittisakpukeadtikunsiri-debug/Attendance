// ============================================================
//  GOOGLE APPS SCRIPT — Face Attendance Backend (v3)
//  ปรับจาก code.gs เดิม: สแกนหลายคนต่อรูป, token, ตรวจ GPS ฝั่งเซิร์ฟเวอร์,
//  หลายไซต์, เข้า/ออกงานอัตโนมัติ, เก็บรูป, ส่งสรุป + ปุ่มยืนยันเข้า LINE
//
//  ต้องตั้งค่า Script Properties (Project Settings > Script properties):
//    API_TOKEN        PIN ทีม (หน้า scan ใช้)
//    ADMIN_TOKEN      PIN แอดมิน (หน้า register ใช้)
//    LINE_TOKEN       Channel access token ของ LINE Messaging API
//    WEBHOOK_KEY      รหัสลับต่อท้าย Webhook URL  (?key=...)
//    CONFIRM_USER_IDS (ไม่บังคับ) userId ที่กดยืนยันได้ คั่นด้วย ,
//
//  Deploy: Deploy > New deployment > Web app
//          Execute as: Me | Who has access: Anyone
// ============================================================

var TZ = 'Asia/Bangkok';
var DEDUP_MINUTES = 3;      // สแกนซ้ำภายในกี่นาทีให้ข้าม
var MAX_NAMES = 40;         // จำนวนชื่อสูงสุดต่อครั้ง
var MAX_PHOTO_CHARS = 3000000;

var ATT_HEADER = ['Timestamp', 'Date', 'Time', 'Name', 'Type', 'SiteId', 'SiteName',
  'Lat', 'Lng', 'DistanceKm', 'MapLink', 'PhotoLink', 'BatchId', 'Status', 'Note'];
var COL_BATCH = 13, COL_STATUS = 14;

// ---------- helpers ----------
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    if (header) s.appendRow(header);
  }
  return s;
}
function clean_(s) {            // กัน formula injection + ตัดความยาว
  return String(s == null ? '' : s).replace(/^[=+\-@\s]+/, '').trim().slice(0, 100);
}
function requireToken_(t, adminOnly) {
  var api = prop_('API_TOKEN'), admin = prop_('ADMIN_TOKEN');
  if (!api || !admin) throw new Error('ยังไม่ได้ตั้งค่า API_TOKEN / ADMIN_TOKEN ใน Script properties');
  t = String(t || '');
  var ok = adminOnly ? (t === admin) : (t === api || t === admin);
  if (!ok) throw new Error('PIN ไม่ถูกต้อง');
}
function distKm_(lat1, lng1, lat2, lng2) {
  var R = 6371, r = Math.PI / 180;
  var dLat = (lat2 - lat1) * r, dLng = (lng2 - lng1) * r;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * r) * Math.cos(lat2 * r) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.asin(Math.sqrt(a));
}
function thaiDate_(d) {
  return Utilities.formatDate(d, TZ, 'dd/MM/') + (Number(Utilities.formatDate(d, TZ, 'yyyy')) + 543);
}

// ---------- รันครั้งเดียวตอนติดตั้ง ----------
function setupSheets() {
  sheet_('Users', ['Name', 'Descriptor', 'RegisteredAt', 'SiteId']);
  var sites = sheet_('Sites', ['SiteId', 'SiteName', 'Lat', 'Lng', 'RadiusKm', 'LineGroupId']);
  if (sites.getLastRow() < 2) sites.appendRow(['site1', 'ไซต์ตัวอย่าง (แก้ไขแถวนี้)', 16.4322, 103.5061, 0.3, '']);
  sheet_('Attendance', ATT_HEADER);
  return 'ok';
}

// ---------- GET ----------
function doGet(e) {
  var p = e.parameter || {};
  try {
    if (p.action === 'getConfig') { requireToken_(p.token, false); return json_(getConfig_()); }
    if (p.action === 'getKnownFaces') { requireToken_(p.token, false); return json_(getKnownFaces_()); }
    return json_({ error: 'Unknown action: ' + p.action });
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

// ---------- POST (หน้าเว็บ + LINE webhook) ----------
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ error: 'Invalid JSON body' }); }

  try {
    if (body.events) {                                   // มาจาก LINE
      var key = (e.parameter || {}).key;
      if (!prop_('WEBHOOK_KEY') || key !== prop_('WEBHOOK_KEY')) return json_({ error: 'forbidden' });
      handleLineEvents_(body.events);
      return json_({ ok: true });
    }
    switch (body.action) {
      case 'registerUser': requireToken_(body.token, true); return json_(registerUser_(body));
      case 'logAttendanceBatch': requireToken_(body.token, false); return json_(logAttendanceBatch_(body));
      default: return json_({ error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

// ---------- Sites / Config ----------
function getSites_() {
  var s = ss_().getSheetByName('Sites');
  if (!s || s.getLastRow() < 2) return [];
  var rows = s.getRange(2, 1, s.getLastRow() - 1, 6).getValues(), out = [];
  rows.forEach(function (r) {
    if (r[0] === '' || r[2] === '' || r[3] === '') return;
    out.push({
      id: String(r[0]), name: String(r[1]), lat: Number(r[2]), lng: Number(r[3]),
      radiusKm: Number(r[4]) || 0.3, lineGroupId: String(r[5] || '')
    });
  });
  return out;
}
function getConfig_() {      // ไม่ส่ง lineGroupId ออกไปหน้าเว็บ
  return { sites: getSites_().map(function (s) {
    return { id: s.id, name: s.name, lat: s.lat, lng: s.lng, radiusKm: s.radiusKm };
  }) };
}
function findSite_(lat, lng) {
  var best = null;
  getSites_().forEach(function (s) {
    var d = distKm_(lat, lng, s.lat, s.lng);
    if (!best || d < best.dist) best = { site: s, dist: d };
  });
  return best;
}

// ---------- Users (ใบหน้า) ----------
function registerUser_(b) {
  var name = clean_(b.name);
  if (!name) throw new Error('กรุณาระบุชื่อ');
  var list = b.descriptors;
  if (!Array.isArray(list) || list.length < 1 || list.length > 10) throw new Error('ข้อมูลใบหน้าไม่ถูกต้อง');
  list.forEach(function (d) {
    if (!Array.isArray(d) || d.length !== 128) throw new Error('ข้อมูลใบหน้าไม่ถูกต้อง (ต้องมี 128 ค่า)');
  });
  var now = new Date(), siteId = clean_(b.siteId);
  var rows = list.map(function (d) { return [name, JSON.stringify(d), now, siteId]; });

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = sheet_('Users', ['Name', 'Descriptor', 'RegisteredAt', 'SiteId']);
    s.getRange(s.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  } finally { lock.releaseLock(); }
  return { success: true, message: 'บันทึกใบหน้า ' + name + ' จำนวน ' + rows.length + ' ภาพ' };
}

function getKnownFaces_() {
  var s = ss_().getSheetByName('Users');
  if (!s || s.getLastRow() < 2) return [];
  var data = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  var map = {}, order = [];
  data.forEach(function (r) {
    var name = String(r[0] || ''), js = r[1];
    if (!name || !js) return;
    try {
      var d = JSON.parse(js);
      if (!map[name]) { map[name] = []; order.push(name); }
      map[name].push(d);
    } catch (e) { Logger.log('descriptor เสีย: ' + name); }
  });
  return order.map(function (n) { return { label: n, descriptors: map[n] }; });
}
function userNameSet_() {
  var s = ss_().getSheetByName('Users'), set = {};
  if (!s || s.getLastRow() < 2) return set;
  s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) set[String(r[0])] = true; });
  return set;
}

// ---------- Attendance (บันทึกทั้งทีมครั้งเดียว) ----------
function logAttendanceBatch_(b) {
  var lat = Number(b.lat), lng = Number(b.lng);
  if (!isFinite(lat) || !isFinite(lng) || (lat === 0 && lng === 0)) throw new Error('ไม่พบตำแหน่ง GPS');
  var names = (Array.isArray(b.names) ? b.names : []).map(clean_).filter(Boolean);
  names = names.filter(function (n, i) { return names.indexOf(n) === i; });
  if (!names.length) throw new Error('ไม่มีรายชื่อ');
  if (names.length > MAX_NAMES) throw new Error('รายชื่อมากเกินไป');
  var force = (b.forceType === 'IN' || b.forceType === 'OUT') ? b.forceType : '';

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // 1) ตรวจรัศมีฝั่งเซิร์ฟเวอร์
    var near = findSite_(lat, lng);
    if (!near) throw new Error('ยังไม่ได้ตั้งค่าไซต์ในชีต Sites');
    if (near.dist > near.site.radiusKm) {
      throw new Error('อยู่นอกรัศมีไซต์ (ใกล้สุด: ' + near.site.name + ' ห่าง ' + near.dist.toFixed(2) + ' กม.)');
    }
    var site = near.site;

    // 2) หาสถานะล่าสุดของแต่ละคนในวันนี้
    var sh = sheet_('Attendance', ATT_HEADER);
    var now = new Date(), todayKey = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    var last = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (!(r[0] instanceof Date)) return;
        if (Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') !== todayKey) return;
        last[String(r[3])] = { type: String(r[4]), ts: r[0] };
      });
    }

    // 3) แยกคนที่บันทึกได้ / ข้าม / ไม่รู้จัก
    var known = userNameSet_(), recorded = [], skipped = [], rejected = [];
    names.forEach(function (n) {
      if (!known[n]) { rejected.push(n); return; }
      var prev = last[n];
      var type = force || (prev && prev.type === 'IN' ? 'OUT' : 'IN');
      if (prev) {
        var mins = (now - prev.ts) / 60000;
        if (mins < DEDUP_MINUTES && (!force || prev.type === type)) { skipped.push(n); return; }
      }
      recorded.push({ name: n, type: type });
    });
    if (!recorded.length) {
      return { success: true, recorded: [], skipped: skipped, rejected: rejected,
        siteName: site.name, message: 'ไม่มีรายการใหม่ให้บันทึก' };
    }

    // 4) เก็บรูป + เขียนทีเดียว
    var batchId = Utilities.formatDate(now, TZ, 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var photoUrl = savePhoto_(b.photoBase64, batchId);
    var dateStr = "'" + thaiDate_(now), timeStr = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    var mapLink = 'https://www.google.com/maps?q=' + lat + ',' + lng;
    var unknown = Math.max(0, Number(b.unknownCount) || 0);
    var note = unknown ? ('ไม่ผ่านการจับคู่ ' + unknown + ' คน') : '';
    var dist = Math.round(near.dist * 1000) / 1000;

    var rows = recorded.map(function (r) {
      return [now, dateStr, timeStr, r.name, r.type, site.id, site.name, lat, lng, dist,
        mapLink, photoUrl, batchId, 'Pending', note];
    });
    var range = sh.getRange(sh.getLastRow() + 1, 1, rows.length, ATT_HEADER.length);
    range.setValues(rows);
    sh.getRange(range.getRow(), 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    // 5) ส่งสรุปเข้า LINE (ล้มเหลวไม่ทำให้ข้อมูลหาย)
    var lineStatus = 'skipped';
    try {
      lineStatus = pushSummary_(site, {
        batchId: batchId, recorded: recorded, skipped: skipped, rejected: rejected,
        unknown: unknown, when: thaiDate_(now) + ' ' + timeStr.slice(0, 5), dist: dist
      });
    } catch (e) { lineStatus = 'failed: ' + e.message; }

    return { success: true, batchId: batchId, siteName: site.name, distanceKm: dist,
      recorded: recorded, skipped: skipped, rejected: rejected, line: lineStatus,
      photo: photoUrl ? 'saved' : 'none' };
  } finally { lock.releaseLock(); }
}

function savePhoto_(b64, batchId) {
  if (!b64) return '';
  b64 = String(b64).replace(/^data:image\/\w+;base64,/, '');
  if (b64.length > MAX_PHOTO_CHARS) return '';
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', batchId + '.jpg');
    var id = prop_('PHOTO_FOLDER_ID'), folder;
    if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
    if (!folder) {
      folder = DriveApp.createFolder('AttendancePhotos');
      PropertiesService.getScriptProperties().setProperty('PHOTO_FOLDER_ID', folder.getId());
    }
    return folder.createFile(blob).getUrl();        // ไฟล์เป็นส่วนตัว เปิดได้เฉพาะเจ้าของ
  } catch (e) {
    Logger.log('savePhoto ล้มเหลว: ' + e.message);
    return '';
  }
}

// ---------- LINE ----------
function linePost_(path, payload) {
  var token = prop_('LINE_TOKEN');
  if (!token) return 0;
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 200) Logger.log('LINE ' + path + ' ' + code + ' ' + res.getContentText());
  return code;
}
function replyLine_(replyToken, text) {
  if (replyToken) linePost_('reply', { replyToken: replyToken, messages: [{ type: 'text', text: text }] });
}

function pushSummary_(site, info) {
  if (!prop_('LINE_TOKEN') || !site.lineGroupId) return 'skipped';
  var ins = [], outs = [];
  info.recorded.forEach(function (r) { (r.type === 'IN' ? ins : outs).push(r.name); });

  var lines = [
    { type: 'text', text: '✅ บันทึกเวลาแล้ว', weight: 'bold', size: 'lg' },
    { type: 'text', text: 'ไซต์: ' + site.name, wrap: true },
    { type: 'text', text: 'เวลา: ' + info.when, wrap: true }
  ];
  if (ins.length) lines.push({ type: 'text', text: 'เข้างาน (' + ins.length + ' คน): ' + ins.join(', '), wrap: true });
  if (outs.length) lines.push({ type: 'text', text: 'เลิกงาน (' + outs.length + ' คน): ' + outs.join(', '), wrap: true });
  if (info.skipped.length) lines.push({ type: 'text', text: 'ข้าม (สแกนซ้ำ): ' + info.skipped.join(', '), wrap: true, size: 'sm', color: '#777777' });
  if (info.rejected.length) lines.push({ type: 'text', text: 'ไม่พบในระบบ: ' + info.rejected.join(', '), wrap: true, size: 'sm', color: '#B3261E' });
  if (info.unknown) lines.push({ type: 'text', text: 'ใบหน้าที่ไม่ผ่านการจับคู่: ' + info.unknown + ' คน', wrap: true, size: 'sm', color: '#B25E00' });

  var msg = {
    type: 'flex',
    altText: 'บันทึกเวลา ' + site.name + ' ' + info.recorded.length + ' คน',
    contents: {
      type: 'bubble',
      body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: lines },
      footer: { type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
        { type: 'button', style: 'primary', height: 'sm',
          action: { type: 'postback', label: 'ถูกต้อง', data: 'confirm=' + info.batchId, displayText: 'ยืนยัน: ถูกต้อง' } },
        { type: 'button', style: 'secondary', height: 'sm',
          action: { type: 'postback', label: 'ขอแก้ไข', data: 'dispute=' + info.batchId, displayText: 'ขอแก้ไขรายการ' } }
      ] }
    }
  };
  var code = linePost_('push', { to: site.lineGroupId, messages: [msg] });
  return code === 200 ? 'sent' : 'failed: HTTP ' + code;
}

function handleLineEvents_(events) {
  events.forEach(function (ev) {
    var src = ev.source || {};
    if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
      if (ev.message.text.trim().toLowerCase() === '!id') {
        replyLine_(ev.replyToken, 'groupId: ' + (src.groupId || '-') + '\nuserId: ' + (src.userId || '-'));
      }
      return;
    }
    if (ev.type === 'postback' && ev.postback && ev.postback.data) {
      var m = /^(confirm|dispute)=(.+)$/.exec(ev.postback.data);
      if (!m) return;
      var allowed = prop_('CONFIRM_USER_IDS').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      if (allowed.length && allowed.indexOf(src.userId) < 0) {
        replyLine_(ev.replyToken, 'ไม่มีสิทธิ์กดยืนยัน');
        return;
      }
      var status = m[1] === 'confirm' ? 'Confirmed' : 'Disputed';
      var n = updateStatus_(m[2], status);
      replyLine_(ev.replyToken, n
        ? (status === 'Confirmed' ? '✔ ยืนยันแล้ว ' + n + ' รายการ' : '⚠ รับเรื่องขอแก้ไข ' + n + ' รายการ — แอดมินจะตรวจสอบ')
        : 'ไม่พบรายการ (อาจถูกลบแล้ว)');
    }
  });
}

function updateStatus_(batchId, status) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName('Attendance');
    if (!sh || sh.getLastRow() < 2) return 0;
    var n = sh.getLastRow() - 1;
    var ids = sh.getRange(2, COL_BATCH, n, 1).getValues();
    var st = sh.getRange(2, COL_STATUS, n, 1).getValues(), count = 0;
    for (var i = 0; i < n; i++) {
      if (String(ids[i][0]) === batchId) { st[i][0] = status; count++; }
    }
    if (count) sh.getRange(2, COL_STATUS, n, 1).setValues(st);
    return count;
  } finally { lock.releaseLock(); }
}
