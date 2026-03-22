// =================================================================================
// GUARD TOUR SYSTEM — Code.gs (PRODUCTION READY v3.0)
// v3.0 SECURITY: ย้าย secrets ทั้งหมดออกจาก hardcode → PropertiesService
//
// ❗ ขั้นตอนตั้งค่าครั้งแรก:
//   1. เปิด Apps Script Editor
//   2. รัน  setupProperties()  เพียงครั้งเดียว — ระบบจะสร้าง Properties ให้ครบ
//   3. แก้ไขค่าจริงใน Project Settings → Script Properties
//   4. Deploy Web App ใหม่ (Execute as: Me, Who has access: Anyone)
// =================================================================================

// =================================================================================
// PROPERTIES HELPER — อ่านค่าจาก Script Properties (ปลอดภัย)
// =================================================================================
function getProp_(key, fallback) {
  try {
    const val = PropertiesService.getScriptProperties().getProperty(key);
    return (val !== null && val !== '') ? val : (fallback || '');
  } catch(e) {
    Logger.log('getProp_ error for ' + key + ': ' + e.message);
    return fallback || '';
  }
}

// CONFIG ที่ไม่ใช่ secret — เก็บไว้ใน code ได้ปลอดภัย
const CONFIG = {
  // ── Non-secret config (ปลอดภัย) ──────────────────────────────────────────────
  GPS_RADIUS_M:        50,
  SESSION_HOURS:       8,
  TIMEZONE:            'Asia/Bangkok',
  DRIVE_FOLDER_NAME:   'GuardTourPhotos',
  REPORT_FOLDER_NAME:  'GuardTourReports',
  LATE_GRACE_MINUTES:  30,
  UNIT_NAME:           'ป.5 พัน.5',

  SHIFTS: [
    { id: 1,  label: 'ผลัดที่ 1',  start: '22:00', end: '00:00' },
    { id: 2,  label: 'ผลัดที่ 2',  start: '00:00', end: '02:00' },
    { id: 3,  label: 'ผลัดที่ 3',  start: '02:00', end: '04:00' },
    { id: 4,  label: 'ผลัดที่ 4',  start: '04:00', end: '06:00' },
    { id: 5,  label: 'ผลัดที่ 5',  start: '06:00', end: '08:00' },
    { id: 6,  label: 'ผลัดที่ 6',  start: '08:00', end: '10:00' },
    { id: 7,  label: 'ผลัดที่ 7',  start: '10:00', end: '12:00' },
    { id: 8,  label: 'ผลัดที่ 8',  start: '12:00', end: '14:00' },
    { id: 9,  label: 'ผลัดที่ 9',  start: '14:00', end: '16:00' },
    { id: 10, label: 'ผลัดที่ 10', start: '16:00', end: '18:00' },
    { id: 11, label: 'ผลัดที่ 11', start: '18:00', end: '20:00' },
    { id: 12, label: 'ผลัดที่ 12', start: '20:00', end: '22:00' },
  ],
};

// ── Secret getters (อ่านจาก Properties ทุกครั้ง) ─────────────────────────────
function SHEET_ID()             { return getProp_('SHEET_ID'); }
function CHANNEL_ACCESS_TOKEN() { return getProp_('CHANNEL_ACCESS_TOKEN'); }
function WEB_APP_URL()          { return getProp_('WEB_APP_URL'); }
function LIFF_ID()              { return getProp_('LIFF_ID'); }
function ADMIN_SECRET()         { return getProp_('ADMIN_SECRET'); }
function SUPERADMIN_SECRET()    { return getProp_('SUPERADMIN_SECRET'); }
function TELEGRAM_BOT_TOKEN()   { return getProp_('TELEGRAM_BOT_TOKEN'); }
function TELEGRAM_CHAT_ID()     { return getProp_('TELEGRAM_CHAT_ID'); }

// =================================================================================
// SETUP — รันครั้งแรกเพื่อสร้าง Script Properties (แก้ค่าจริงใน Project Settings)
// =================================================================================
function setupProperties() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperties();

  // ตั้งค่า default เฉพาะที่ยังไม่มี — ไม่ทับค่าที่ตั้งแล้ว
  const defaults = {
    'SHEET_ID':              'ใส่ Google Sheet ID ที่นี่',
    'CHANNEL_ACCESS_TOKEN':  'ใส่ LINE Channel Access Token ที่นี่',
    'WEB_APP_URL':           'ใส่ Web App URL ที่นี่',
    'LIFF_ID':               'ใส่ LIFF ID ที่นี่',
    'ADMIN_SECRET':          'GUARD2025',
    'SUPERADMIN_SECRET':     'SUPER2025',
    'TELEGRAM_BOT_TOKEN':    'ใส่ Telegram Bot Token ที่นี่',
    'TELEGRAM_CHAT_ID':      'ใส่ Telegram Chat ID ที่นี่',
  };

  let added = 0;
  Object.entries(defaults).forEach(([k, v]) => {
    if (!existing[k]) {
      props.setProperty(k, v);
      added++;
    }
  });

  Logger.log(
    '✅ setupProperties เสร็จแล้ว\n' +
    'เพิ่ม ' + added + ' property ใหม่\n' +
    'ไปที่ Project Settings → Script Properties เพื่อแก้ค่าจริง\n' +
    'Keys ที่ต้องแก้: SHEET_ID, CHANNEL_ACCESS_TOKEN, WEB_APP_URL, LIFF_ID, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID'
  );
}

/** ดูค่า Properties ทั้งหมดที่ตั้งไว้ (ไม่แสดงค่าของ token ยาว) */
function showProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const sensitive = ['CHANNEL_ACCESS_TOKEN', 'TELEGRAM_BOT_TOKEN'];
  const lines = Object.entries(props).map(([k, v]) => {
    const display = sensitive.includes(k)
      ? v.substring(0, 6) + '...' + v.substring(v.length - 4) + ' (length: ' + v.length + ')'
      : v;
    return k + ' = ' + display;
  });
  Logger.log('=== Script Properties ===\n' + lines.join('\n'));
}

// =================================================================================
// SHEET HELPERS
// =================================================================================
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID());
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a1a2e').setFontColor('#f59e0b').setFontWeight('bold');
  }
  return sheet;
}

function guardLogSheet() {
  return getSheet_('AdminLog', [
    'UserId','LoginTimestamp','UserName','Status','AdminType','LogoutTimestamp'
  ]);
}

function tourSheet() {
  return getSheet_('GuardTours', [
    'TourId','CheckpointId','CheckpointName','GuardName','GuardUserId',
    'Timestamp','Status','Notes','PhotoUrl','TourStartTime','CheckIn_Lat','CheckIn_Lng'
  ]);
}

function checkpointSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID());
  let sheet = ss.getSheetByName('Checkpoints');
  if (!sheet) {
    sheet = ss.insertSheet('Checkpoints');
    const headers = ['CheckpointId','CheckpointName','Location','Lat','Lng','Active','SortOrder'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a1a2e').setFontColor('#f59e0b').setFontWeight('bold');
    // ข้อมูล default จุดตรวจ 8 จุด (ซิงค์กับ qr-checkpoints.html)
    sheet.getRange(2, 1, 7, 7).setValues([
      ['CP-01','ประตูทิศเหนือ','ทางเข้าด้านเหนือ',       0, 0, 'TRUE', 1],
      ['CP-02','ประตูทิศใต้',  'ทางเข้าด้านใต้',         0, 0, 'TRUE', 2],
      ['CP-03','คลัง สป.5',    'คลังสิ่งปลูกสร้าง',      0, 0, 'TRUE', 3],
      ['CP-04','โรงปืนร้อย.1', 'โรงเก็บอาวุธ ร้อย 1',   0, 0, 'TRUE', 4],
      ['CP-05','โรงน้ำบิ๊กกัน','อาคารระบบน้ำ',           0, 0, 'TRUE', 5],
      ['CP-06','บ้านพักผู้พัน','บ้านพักนายทหาร',         0, 0, 'TRUE', 6],
      ['CP-07','บ้านพักซอย 5', 'บ้านพักซอย 5',           0, 0, 'TRUE', 7],
    ]);
    Logger.log('✅ สร้าง Checkpoints sheet พร้อมข้อมูล 8 จุด');
  }
  return sheet;
}

function lateAlertLogSheet() {
  return getSheet_('LateAlertLog', [
    'AlertKey','SentAt','ShiftId','ShiftLabel','Message'
  ]);
}

// =================================================================================
// HTTP HANDLERS
// =================================================================================
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  if (p.page === 'debug') return HtmlService.createHtmlOutput(getDebugInfo());

  const action = p.action;

  if (action === 'check_admin')     return jsonResp(checkAdminForLiff(p.userId));
  if (action === 'get_checkpoints') return jsonResp(getCheckpoints());
  if (action === 'get_stats')       return jsonResp(getTodayStats());

  // Dashboard secret verification (ไม่เปิดเผย secret ใน response)
  if (action === 'verify_dashboard_secret') {
    const secret = p.secret || '';
    const isValid = (secret !== '' && secret === SUPERADMIN_SECRET());
    return jsonResp({ ok: isValid });
  }

  // Admin Panel — Checkpoint Management
  if (action === 'get_checkpoints_admin') return jsonResp(getCheckpointsAdmin());
  if (action === 'add_checkpoint')        return jsonResp(addCheckpoint(p));
  if (action === 'update_checkpoint')     return jsonResp(updateCheckpoint(p));
  if (action === 'delete_checkpoint')     return jsonResp(deleteCheckpoint(p.checkpointId));

  // Dashboard API
  if (action === 'get_tour_data')           return jsonResp(getTourData(p));
  if (action === 'get_active_guards')       return jsonResp(getActiveGuards());
  if (action === 'get_dashboard_summary')   return jsonResp(getDashboardSummary());
  if (action === 'export_excel')            return exportExcel(p);

  // Report API
  if (action === 'generate_daily_report')   return jsonResp(generateDailyReport(p.date));
  if (action === 'generate_monthly_report') return jsonResp(generateMonthlyReport(p.year, p.month));
  if (action === 'get_stats_30days')        return jsonResp(getStats30Days());
  if (action === 'get_gps_tracks')          return jsonResp(getGpsTracks(p));

  if (action === 'guard_checkin') {
    try {
      return handleGuardCheckin(buildCheckinFromParams_(p));
    } catch(err) {
      Logger.log('GET guard_checkin error: ' + err.message);
      return jsonResp({ status: 'error', message: err.message });
    }
  }

  if (action === 'guard_tour_complete') {
    try {
      let summaryArr = [];
      if (p.summary) { try { summaryArr = JSON.parse(p.summary); } catch(_) {} }
      return handleGuardTourComplete({
        tourId: p.tourId || '', userId: p.userId || '',
        guardName: p.guardName || '', tourStartTime: p.tourStartTime || '',
        tourEndTime: p.tourEndTime || new Date().toISOString(),
        summary: summaryArr,
      });
    } catch(err) {
      return jsonResp({ status: 'error', message: err.message });
    }
  }

  return HtmlService.createHtmlOutput(
    '<h2>🛡️ Guard Tour System — Online ✅</h2><p>Version 3.0</p>'
  );
}

function buildCheckinFromParams_(p) {
  return {
    action: p.action,
    tourId: p.tourId || '', userId: p.userId || '',
    guardName: p.guardName || '', checkpointId: p.checkpointId || '',
    checkpointName: p.checkpointName || '', status: p.status || 'normal',
    notes: p.notes || '', timestamp: p.timestamp || new Date().toISOString(),
    tourStartTime: p.tourStartTime || '', gpsLat: p.gpsLat || '',
    gpsLng: p.gpsLng || '', imageBase64: null, imageMime: 'image/jpeg',
  };
}

function doPost(e) {
  try {
    let data;
    const p = (e && e.parameter) ? e.parameter : {};

    if (e.postData && e.postData.contents) {
      try {
        const parsed = JSON.parse(e.postData.contents);
        if (parsed.action === 'upload_and_notify_photo') return handleUploadAndNotifyPhoto(parsed);
        if (parsed.action === 'upload_photo')            return handleUploadPhoto(parsed);
        if (parsed.action === 'guard_checkin')           return handleGuardCheckin(parsed);
        if (parsed.action === 'guard_tour_complete') {
          if (typeof parsed.summary === 'string') {
            try { parsed.summary = JSON.parse(parsed.summary); } catch(_) { parsed.summary = []; }
          }
          return handleGuardTourComplete(parsed);
        }
        data = parsed;
      } catch(je) { Logger.log('JSON parse error: ' + je.message); }
    }

    if (!data && p.action) {
      if (p.action === 'guard_checkin')           { data = buildCheckinFromParams_(p); }
      else if (p.action === 'upload_photo')        { return handleUploadPhoto(p); }
      else if (p.action === 'upload_and_notify_photo') { return handleUploadAndNotifyPhoto(p); }
      else if (p.action === 'guard_tour_complete') {
        data = Object.assign({}, p);
        if (typeof data.summary === 'string') {
          try { data.summary = JSON.parse(data.summary); } catch(_) { data.summary = []; }
        }
      } else if (p.payload) {
        try { data = JSON.parse(p.payload); } catch(_) {}
      }
    }

    if (!data) return jsonResp({ status: 'error', message: 'ไม่พบข้อมูลที่ส่งมา' });

    if (data.action === 'guard_checkin')           return handleGuardCheckin(data);
    if (data.action === 'guard_tour_complete')     return handleGuardTourComplete(data);
    if (data.action === 'upload_photo')            return handleUploadPhoto(data);
    if (data.action === 'upload_and_notify_photo') return handleUploadAndNotifyPhoto(data);

    if (data.events && data.events.length > 0) {
      data.events.forEach(ev => {
        try { handleWebhook(ev); } catch(err) { Logger.log('webhook err: ' + err.message); }
      });
    }
    return jsonResp({ status: 'ok' });
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResp({ status: 'error', message: err.message });
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =================================================================================
// AUTH
// =================================================================================
function checkAdminForLiff(userId) {
  if (!userId) return { isAdmin: false, reason: 'ไม่พบข้อมูลผู้ใช้' };
  try {
    const data = guardLogSheet().getDataRange().getValues();
    const now  = new Date();
    const SESSION_MS = CONFIG.SESSION_HOURS * 60 * 60 * 1000;

    for (let i = data.length - 1; i > 0; i--) {
      const [uid, loginTime, userName, status, adminType] = data[i];
      if (uid !== userId || status !== 'ACTIVE') continue;
      if (adminType === 'GUARD') {
        const elapsed = now - new Date(loginTime);
        if (elapsed > SESSION_MS) {
          guardLogSheet().getRange(i + 1, 4).setValue('EXPIRED');
          return { isAdmin: false, reason: `เซสชันหมดอายุ (${CONFIG.SESSION_HOURS} ชั่วโมง)\nกรุณาพิมพ์รหัสลับใน LINE Chat ใหม่` };
        }
      }
      return { isAdmin: true, adminType, userName };
    }
    return { isAdmin: false, reason: 'คุณยังไม่ได้ Login\nพิมพ์รหัสลับใน LINE Chat ก่อนเปิดแอป' };
  } catch(e) {
    Logger.log('checkAdminForLiff error: ' + e.message);
    return { isAdmin: false, reason: 'ระบบขัดข้อง: ' + e.message };
  }
}

// =================================================================================
// CHECKPOINTS
// =================================================================================
function getCheckpoints() {
  try {
    const data = checkpointSheet().getDataRange().getValues();
    if (data.length <= 1) return { status: 'ok', checkpoints: [] };
    const cps = data.slice(1)
      .filter(r => r[0] && String(r[5]).toUpperCase() !== 'FALSE')
      .sort((a, b) => (Number(a[6]) || 99) - (Number(b[6]) || 99))
      .map(r => ({
        id:       String(r[0]).trim(),
        name:     String(r[1]).trim(),
        location: String(r[2]).trim(),
        lat:      parseFloat(r[3]) || 0,
        lng:      parseFloat(r[4]) || 0,
        icon:     '📍',
      }));
    return { status: 'ok', checkpoints: cps };
  } catch(e) {
    Logger.log('getCheckpoints error: ' + e.message);
    return { status: 'error', message: e.message, checkpoints: [] };
  }
}

// =================================================================================
// UPLOAD PHOTO
// =================================================================================
function handleUploadPhoto(p) {
  try {
    const tourId       = p.tourId       || '';
    const checkpointId = p.checkpointId || '';
    const imageBase64  = p.imageBase64  || '';
    const imageMime    = p.imageMime    || 'image/jpeg';
    const timestamp    = p.timestamp    || '';

    if (!imageBase64 || !tourId || !checkpointId) {
      return jsonResp({ status: 'ok', photoUrl: '', skipped: true });
    }

    const photoUrl = savePhotoToDrive(imageBase64, imageMime, tourId, checkpointId);

    if (photoUrl) {
      try {
        const sheet = tourSheet(), data = sheet.getDataRange().getValues();
        const tsNorm = timestamp ? formatISOFromSheet_(new Date(timestamp.replace(' ','T'))).substring(0, 16) : '';
        for (let i = data.length - 1; i > 0; i--) {
          if (String(data[i][0]) !== tourId || String(data[i][1]) !== checkpointId) continue;
          if (tsNorm) {
            const rowNorm = data[i][5] ? formatISOFromSheet_(data[i][5]).substring(0, 16) : '';
            if (rowNorm !== tsNorm) continue;
          }
          sheet.getRange(i + 1, 9).setValue(photoUrl);
          break;
        }
      } catch(updateErr) { Logger.log('updatePhotoUrl error: ' + updateErr.message); }
    }
    return jsonResp({ status: 'ok', photoUrl });
  } catch(e) {
    Logger.log('handleUploadPhoto error: ' + e.message);
    return jsonResp({ status: 'ok', photoUrl: '', error: e.message });
  }
}

function handleUploadAndNotifyPhoto(p) {
  try {
    const tourId       = p.tourId       || '';
    const checkpointId = p.checkpointId || '';
    const imageBase64  = p.imageBase64  || '';
    const imageMime    = p.imageMime    || 'image/jpeg';
    const timestamp    = p.timestamp    || '';

    Logger.log('handleUploadAndNotifyPhoto: tourId=' + tourId + ' cpId=' + checkpointId + ' b64len=' + imageBase64.length);

    if (!imageBase64 || !tourId || !checkpointId) {
      return jsonResp({ status: 'ok', photoUrl: '', skipped: true });
    }

    const photoUrl = savePhotoToDrive(imageBase64, imageMime, tourId, checkpointId);

    if (photoUrl) {
      try {
        const sheet = tourSheet(), data = sheet.getDataRange().getValues();
        const tsNorm = timestamp ? formatISOFromSheet_(new Date(timestamp.replace(' ','T'))).substring(0, 16) : '';
        for (let i = data.length - 1; i > 0; i--) {
          if (String(data[i][0]) !== tourId || String(data[i][1]) !== checkpointId) continue;
          if (tsNorm) {
            const rowNorm = data[i][5] ? formatISOFromSheet_(data[i][5]).substring(0, 16) : '';
            if (rowNorm !== tsNorm) continue;
          }
          sheet.getRange(i + 1, 9).setValue(photoUrl);
          break;
        }
      } catch(updateErr) { Logger.log('updatePhotoUrl in notify error: ' + updateErr.message); }
      try { pushImageToAllSuperAdmins(photoUrl); } catch(lineErr) { Logger.log('pushImage error: ' + lineErr.message); }
    }
    return jsonResp({ status: 'ok', photoUrl });
  } catch(e) {
    Logger.log('handleUploadAndNotifyPhoto error: ' + e.message);
    return jsonResp({ status: 'ok', photoUrl: '', error: e.message });
  }
}

// =================================================================================
// GUARD CHECKIN
// =================================================================================
function handleGuardCheckin(data) {
  try {
    const { tourId, userId, guardName, checkpointId, checkpointName, status, notes, timestamp, tourStartTime, gpsLat, gpsLng } = data;
    if (!tourId || !userId || !checkpointId) return jsonResp({ status: 'error', message: 'ข้อมูลไม่ครบถ้วน (tourId/userId/checkpointId)' });

    const fmt = d => {
      try {
        const parsed = new Date(d);
        if (isNaN(parsed.getTime())) return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
        return Utilities.formatDate(parsed, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      } catch(_) {
        return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      }
    };

    const nowIso         = new Date().toISOString();
    const tsFormatted    = fmt(timestamp     || nowIso);
    const startFormatted = fmt(tourStartTime || timestamp || nowIso);
    const latVal = gpsLat ? parseFloat(String(gpsLat)) : NaN;
    const lngVal = gpsLng ? parseFloat(String(gpsLng)) : NaN;
    const isAbnormal = (String(status || '').toLowerCase() === 'abnormal' || status === 'ผิดปกติ');

    tourSheet().appendRow([
      tourId, checkpointId, checkpointName || checkpointId,
      guardName || '', userId,
      tsFormatted, isAbnormal ? 'ผิดปกติ' : 'ปกติ',
      notes || '', '',
      startFormatted,
      !isNaN(latVal) ? latVal.toFixed(6) : '',
      !isNaN(lngVal) ? lngVal.toFixed(6) : '',
    ]);

    notifyCheckin({ userId, guardName, checkpointId, checkpointName, status: isAbnormal ? 'abnormal' : 'normal', notes, tourStartTime: startFormatted });
    return jsonResp({ status: 'ok', saved: true, timestamp: tsFormatted });
  } catch(e) {
    Logger.log('handleGuardCheckin error: ' + e.message + '\n' + e.stack);
    return jsonResp({ status: 'error', message: e.message });
  }
}

// =================================================================================
// HELPER — วันที่ไทย / ผลัดเวร
// =================================================================================
function thaiDateShort(d) {
  const MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear() + 543}`;
}

function thaiDateFull(d) {
  const MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return `${d.getDate()} ${MONTHS[d.getMonth()]} พ.ศ. ${d.getFullYear()+543}`;
}

function detectShift_(startDate) {
  try {
    const hhmm = Utilities.formatDate(startDate, CONFIG.TIMEZONE, 'HH:mm');
    const [hh, mm] = hhmm.split(':').map(Number);
    const startMin = hh * 60 + mm;
    for (const shObj of CONFIG.SHIFTS) {
      const [shH, shM] = shObj.start.split(':').map(Number);
      const [eH, eM]   = shObj.end.split(':').map(Number);
      let s = shH * 60 + shM, e = eH * 60 + eM;
      if (e <= s) e += 24 * 60;
      const adj = startMin + (startMin < s && s >= 12 * 60 ? 24 * 60 : 0);
      if (adj >= s && adj < e) return shObj;
    }
    return null;
  } catch(_) { return null; }
}

function shiftRangeStr_(startDate, endDate) {
  const sh = detectShift_(startDate);
  if (sh) return `${sh.start.replace(':','.')}-${sh.end.replace(':','.')} (${sh.label})`;
  const fd = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
  return endDate ? `${fd(startDate)}-${fd(endDate)}` : fd(startDate);
}

function formatISOFromSheet_(val) {
  try {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    return String(val).trim();
  } catch(_) { return ''; }
}

// =================================================================================
// NOTIFY CHECKIN
// =================================================================================
function notifyCheckin({ userId, guardName, checkpointId, checkpointName, status, notes, tourStartTime }) {
  try {
    const isAbnormal = (status === 'abnormal' || status === 'ผิดปกติ');

    // ปกติ → ไม่ส่ง Telegram ทันที (รอสรุปรอบ handleGuardTourComplete แทน)
    // ผิดปกติ → ส่งแจ้งเตือนด่วนทันที
    if (!isAbnormal) return;

    const now = new Date();
    const startDate = tourStartTime
      ? new Date(String(tourStartTime).replace(' ', 'T'))
      : now;
    let displayName = guardName || '—';
    try { if (userId) { const p = getUserProfile(userId); if (p && p.displayName) displayName = p.displayName; } } catch(_) {}
    let nowDate = '', nowTimeDot = '';
    try {
      nowDate    = thaiDateShort(now);
      nowTimeDot = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
    } catch(_) {
      nowDate    = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yy');
      nowTimeDot = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
    }
    const shiftStr = shiftRangeStr_(startDate);
    const msg =
      `🚨 พบความผิดปกติ!\n` +
      `━━━━━━━━━━━━━━━━\n` +
      `📅 ${nowDate}, ${nowTimeDot}\n` +
      `⏰ เวลา ${shiftStr}\n` +
      `📍 จุด: ${checkpointName} (${checkpointId})\n` +
      `📝 รายละเอียด: ${notes || '(ไม่ระบุ)'}\n` +
      `━━━━━━━━━━━━━━━━\n` +
      `👤 ${displayName}\n` +
      `📷 รูปภาพจะถูกส่งต่อในอีกสักครู่`;
    pushToAllSuperAdmins(msg);
  } catch(e) { Logger.log('notifyCheckin error: ' + e.message); }
}

// =================================================================================
// TOUR COMPLETE
// =================================================================================
function handleGuardTourComplete(data) {
  try {
    const { userId, tourStartTime, tourEndTime } = data;
    let displayName = data.guardName || '—';
    try { if (userId) { const p = getUserProfile(userId); if (p && p.displayName) displayName = p.displayName; } } catch(_) {}
    let summaryArr = data.summary || [];
    if (typeof summaryArr === 'string') { try { summaryArr = JSON.parse(summaryArr); } catch(_) { summaryArr = []; } }

    const now = new Date();
    const parseTs = ts => {
      if (!ts) return now;
      const d = new Date(String(ts).replace(' ', 'T'));
      return isNaN(d.getTime()) ? now : d;
    };
    const start    = parseTs(tourStartTime);
    const end      = parseTs(tourEndTime);
    const duration = Math.round((end - start) / 60000);

    const abnormal = summaryArr.filter(s =>
      String(s.status || '').toLowerCase() === 'abnormal' || s.status === 'ผิดปกติ'
    );
    const normal = summaryArr.filter(s =>
      String(s.status || '').toLowerCase() !== 'abnormal' && s.status !== 'ผิดปกติ'
    );

    let nowDate = '', nowTimeDot = '';
    try {
      nowDate    = thaiDateShort(now);
      nowTimeDot = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
    } catch(_) {
      nowDate    = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yy');
      nowTimeDot = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
    }

    const shiftStr = shiftRangeStr_(start, end);
    const startStr = Utilities.formatDate(start, CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');
    const endStr   = Utilities.formatDate(end,   CONFIG.TIMEZONE, 'HH:mm').replace(':', '.');

    // รายการจุดตรวจ — แสดงผิดปกติก่อน ตามด้วยปกติ
    const cpLines = [
      ...abnormal.map((s, i) =>
        `⚠️ ${s.checkpointName}${s.notes ? ' — ' + s.notes : ''}`
      ),
      ...normal.map((s, i) =>
        `✅ ${s.checkpointName}`
      ),
    ].join('\n');

    let msg;
    if (abnormal.length === 0) {
      msg =
        `เมื่อ  ${nowDate}, ${nowTimeDot}\n` +
        `เวลา ${shiftStr}\n` +
        `ได้ตรวจเวรยามจุดเสี่ยงจุดล่อแหลมภายในหน่วยฯ` +
        ` ผลการปฏิบัติเป็นไปด้วยความเรียบร้อย เหตุการณ์ทั่วไปปกติครับ\n` +
        `━━━━━━━━━━━━━━━━\n` +
        `📍 จุดตรวจ (${summaryArr.length} จุด):\n${cpLines}\n` +
        `━━━━━━━━━━━━━━━━\n` +
        `⏱️ ${startStr}–${endStr} (${duration} นาที)  👤 ${displayName} ✔️`;
    } else {
      msg =
        `เมื่อ  ${nowDate}, ${nowTimeDot}\n` +
        `เวลา ${shiftStr}\n` +
        `ได้ตรวจเวรยามจุดเสี่ยงจุดล่อแหลมภายในหน่วยฯ` +
        ` พบความผิดปกติ ${abnormal.length} จุด ดังนี้\n` +
        `━━━━━━━━━━━━━━━━\n` +
        `📍 รายงานจุดตรวจทั้งหมด (${summaryArr.length} จุด):\n${cpLines}\n` +
        `━━━━━━━━━━━━━━━━\n` +
        `⏱️ ${startStr}–${endStr} (${duration} นาที)  👤 ${displayName} ✔️`;
    }
    pushToAllSuperAdmins(msg);
    return jsonResp({ status: 'ok' });
  } catch(e) { Logger.log('handleGuardTourComplete error: ' + e.message); return jsonResp({ status: 'error', message: e.message }); }
}

// =================================================================================
// LINE WEBHOOK
// =================================================================================
function handleWebhook(event) {
  const replyToken = event.replyToken, userId = event.source && event.source.userId;
  if (!replyToken || !userId) return;
  if (event.type !== 'message' || !event.message || event.message.type !== 'text') return;
  const text = (event.message.text || '').trim();
  const cmds = {
    [ADMIN_SECRET()]:      () => loginGuard(replyToken, userId, 'GUARD'),
    [SUPERADMIN_SECRET()]: () => loginGuard(replyToken, userId, 'SUPER_ADMIN'),
    'logout':        () => logoutGuard(replyToken, userId),
    'ออกจากระบบ':   () => logoutGuard(replyToken, userId),
    'รายงานวันนี้': () => sendTodayReport(replyToken, userId),
    'report':        () => sendTodayReport(replyToken, userId),
    'สถานะ':        () => sendSystemStatus(replyToken),
    'status':        () => sendSystemStatus(replyToken),
    'help':          () => sendHelp(replyToken),
    'ช่วยเหลือ':    () => sendHelp(replyToken),
  };
  if (cmds[text]) { cmds[text](); return; }
  sendMainMenu(replyToken, userId);
}

// =================================================================================
// LOGIN / LOGOUT
// =================================================================================
function loginGuard(replyToken, userId, adminType) {
  try {
    const sheet = guardLogSheet(), data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId && data[i][3] === 'ACTIVE') {
        sheet.getRange(i + 1, 4).setValue('INACTIVE');
        sheet.getRange(i + 1, 6).setValue(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));
      }
    }
    const profile  = getUserProfile(userId);
    const userName = (profile && profile.displayName) || 'ยาม';
    const loginTs  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([userId, loginTs, userName, 'ACTIVE', adminType, '']);
    const liffUrl    = `https://liff.line.me/${LIFF_ID()}`;
    const roleText   = adminType === 'SUPER_ADMIN' ? '👑 Supervisor' : '🛡️ ยาม';
    const sessionText = adminType === 'GUARD' ? `${CONFIG.SESSION_HOURS} ชั่วโมง` : 'ไม่มีหมดอายุ';
    replyMessageAdvanced(replyToken, [{ type: 'text', text: `✅ เข้าระบบสำเร็จ\n${roleText}: ${userName}\nSession: ${sessionText}`,
      quickReply: { items: [
        { type:'action', action:{ type:'uri',     label:'🚀 เปิดแอปตรวจรอบ', uri: liffUrl }},
        { type:'action', action:{ type:'message', label:'📊 รายงานวันนี้',    text:'รายงานวันนี้' }},
        { type:'action', action:{ type:'message', label:'📡 สถานะระบบ',       text:'สถานะ' }},
        { type:'action', action:{ type:'message', label:'🚪 ออกจากระบบ',      text:'logout' }},
      ]}
    }]);
  } catch(e) { Logger.log('loginGuard error: ' + e.message); replyMessage(replyToken, '❌ Login ผิดพลาด: ' + e.message); }
}

function logoutGuard(replyToken, userId) {
  const sheet = guardLogSheet(), data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId && data[i][3] === 'ACTIVE') {
      sheet.getRange(i + 1, 4).setValue('INACTIVE');
      sheet.getRange(i + 1, 6).setValue(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));
      found = true;
    }
  }
  replyMessageAdvanced(replyToken, [{ type: 'text', text: found ? '👋 ออกจากระบบเรียบร้อยแล้ว' : '⚠️ ไม่พบ session ที่ active อยู่',
    quickReply: { items: [
      { type:'action', action:{ type:'message', label:'🔑 Login ใหม่ (ยาม)',        text: ADMIN_SECRET() }},
      { type:'action', action:{ type:'message', label:'🔑 Login ใหม่ (Supervisor)', text: SUPERADMIN_SECRET() }},
    ]}
  }]);
}

// =================================================================================
// REPORTS (LINE)
// =================================================================================
function getTodayStats() {
  try {
    const data = tourSheet().getDataRange().getValues();
    const todayTH = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const today = data.slice(1).filter(r => {
      if (!r[5]) return false;
      try { return Utilities.formatDate(new Date(r[5]), CONFIG.TIMEZONE, 'dd/MM/yyyy') === todayTH; } catch(_) { return false; }
    });
    return { status: 'ok', date: todayTH, totalCheckins: today.length, totalTours: [...new Set(today.map(r => r[0]))].length, totalAbnormal: today.filter(r => r[6] === 'ผิดปกติ').length };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function sendTodayReport(replyToken, userId) {
  const auth = checkAdminForLiff(userId);
  if (!auth.isAdmin) { replyMessage(replyToken, '❌ กรุณา Login ก่อนดูรายงาน'); return; }
  const data = tourSheet().getDataRange().getValues();
  const todayTH = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const todayRows = data.slice(1).filter(r => {
    if (!r[5]) return false;
    try { return Utilities.formatDate(new Date(r[5]), CONFIG.TIMEZONE, 'dd/MM/yyyy') === todayTH; } catch(_) { return false; }
  });
  if (todayRows.length === 0) {
    replyMessageAdvanced(replyToken, [{ type: 'text', text: `📋 ยังไม่มีการตรวจรอบวันนี้ (${todayTH})`,
      quickReply: { items: [{ type:'action', action:{ type:'uri', label:'🚀 เปิดแอปตรวจรอบ', uri:`https://liff.line.me/${LIFF_ID()}` }}] }
    }]); return;
  }
  const tours = {};
  todayRows.forEach(r => { if (!tours[r[0]]) tours[r[0]] = { guard: r[3], start: r[9], rows: [] }; tours[r[0]].rows.push(r); });
  let msg = `📊 รายงานตรวจรอบ\nวันที่: ${todayTH}\n━━━━━━━━━━━━━━━━\n📦 รวม ${Object.keys(tours).length} รอบ | ${todayRows.length} จุด\n`;
  Object.entries(tours).forEach(([tourId, t]) => {
    const ab = t.rows.filter(r => r[6] === 'ผิดปกติ');
    const startTime = t.start ? Utilities.formatDate(new Date(t.start), CONFIG.TIMEZONE, 'HH:mm') : '—';
    msg += `\n🔹 ${tourId}\n   👤 ${t.guard} | เริ่ม ${startTime}\n   📍 ${t.rows.length} จุด | ` +
      (ab.length > 0 ? `⚠️ ผิดปกติ ${ab.length} จุด\n   ${ab.map(r=>r[2]).join(', ')}` : '✅ ทุกจุดปกติ') + '\n';
  });
  replyMessageAdvanced(replyToken, [{ type: 'text', text: msg,
    quickReply: { items: [
      { type:'action', action:{ type:'message', label:'📡 สถานะระบบ',       text:'สถานะ' }},
      { type:'action', action:{ type:'uri',     label:'🚀 เปิดแอปตรวจรอบ', uri:`https://liff.line.me/${LIFF_ID()}` }},
      { type:'action', action:{ type:'message', label:'🚪 ออกจากระบบ',     text:'logout' }},
    ]}
  }]);
}

function sendSystemStatus(replyToken) {
  const tourData = tourSheet().getDataRange().getValues();
  const todayTH  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const todayCount = tourData.slice(1).filter(r => { try { return Utilities.formatDate(new Date(r[5]), CONFIG.TIMEZONE, 'dd/MM/yyyy') === todayTH; } catch(_) { return false; } }).length;
  const admins = guardLogSheet().getDataRange().getValues();
  const activeGuards      = admins.slice(1).filter(r => r[3]==='ACTIVE' && r[4]==='GUARD').map(r => r[2]);
  const activeSuperAdmins = admins.slice(1).filter(r => r[3]==='ACTIVE' && r[4]==='SUPER_ADMIN').map(r => r[2]);
  replyMessageAdvanced(replyToken, [{ type: 'text',
    text: `📊 สถานะระบบตรวจรอบ\n━━━━━━━━━━━━━━━━\n📅 วันนี้ตรวจ: ${todayCount} รายการ\n📦 ทั้งหมด: ${tourData.length-1} รายการ\n⚠️ ผิดปกติสะสม: ${tourData.slice(1).filter(r=>r[6]==='ผิดปกติ').length} จุด\n🛡️ ยามออนไลน์: ${activeGuards.join(', ') || 'ไม่มี'}\n👑 Supervisor: ${activeSuperAdmins.join(', ') || 'ไม่มี'}`,
    quickReply: { items: [
      { type:'action', action:{ type:'message', label:'📊 รายงานวันนี้',    text:'รายงานวันนี้' }},
      { type:'action', action:{ type:'uri',     label:'🚀 เปิดแอปตรวจรอบ', uri:`https://liff.line.me/${LIFF_ID()}` }},
      { type:'action', action:{ type:'message', label:'🚪 ออกจากระบบ',     text:'logout' }},
    ]}
  }]);
}

function sendHelp(replyToken) {
  replyMessage(replyToken, `📖 คำสั่งที่ใช้ได้\n━━━━━━━━━━━━━━━━\n🔑 รหัสยาม — Login เป็นยาม\n🔑 รหัส Supervisor — Login เป็น Supervisor\n📊 รายงานวันนี้ — ดูรายงานประจำวัน\n📡 สถานะ — ดูสถานะระบบ\n🚪 logout — ออกจากระบบ\n━━━━━━━━━━━━━━━━\nⓘ ยาม: Session ${CONFIG.SESSION_HOURS}ชม.\nSupervisor: ไม่หมดอายุ`);
}

function sendMainMenu(replyToken, userId) {
  const auth = checkAdminForLiff(userId);
  if (!auth.isAdmin) {
    replyMessageAdvanced(replyToken, [{ type: 'text', text: '🛡️ Guard Tour System\n\nพิมพ์รหัสลับเพื่อเข้าระบบ\nหรือพิมพ์ "ช่วยเหลือ" เพื่อดูคำสั่ง',
      quickReply: { items: [
        { type:'action', action:{ type:'message', label:'❓ ช่วยเหลือ', text: 'help' }},
      ]}
    }]); return;
  }
  replyMessageAdvanced(replyToken, [{ type: 'text',
    text: `🛡️ สวัสดี ${auth.userName}\nสถานะ: ${auth.adminType==='SUPER_ADMIN'?'👑 Supervisor':'🛡️ ยาม'}\n\nเลือกรายการที่ต้องการ:`,
    quickReply: { items: [
      { type:'action', action:{ type:'uri',     label:'🚀 เปิดแอปตรวจรอบ', uri:`https://liff.line.me/${LIFF_ID()}` }},
      { type:'action', action:{ type:'message', label:'📊 รายงานวันนี้',    text:'รายงานวันนี้' }},
      { type:'action', action:{ type:'message', label:'📡 สถานะระบบ',       text:'สถานะ' }},
      { type:'action', action:{ type:'message', label:'🚪 ออกจากระบบ',      text:'logout' }},
    ]}
  }]);
}

// =================================================================================
// NOTIFY ABNORMAL (legacy)
// =================================================================================
function notifyAbnormal({ tourId, guardName, checkpointName, checkpointId, notes, photoUrl, timestamp }) {
  try {
    const time = Utilities.formatDate(new Date(timestamp || Date.now()), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
    pushToAllSuperAdmins(`🚨 พบความผิดปกติ!\n━━━━━━━━━━━━━━━━\n📍 จุด: ${checkpointName} (${checkpointId})\n👤 ยาม: ${guardName}\n🕐 เวลา: ${time}\n📝 รายละเอียด: ${notes || '(ไม่ระบุ)'}\n${photoUrl ? '🖼️ ดูรูป: ' + photoUrl : '📷 รูปจะอัปโหลดในอีกสักครู่'}`);
  } catch(e) { Logger.log('notifyAbnormal error: ' + e.message); }
}

function pushToAllSuperAdmins(message) { telegramSendMessage(message); }

function pushImageToAllSuperAdmins(driveUrl) {
  if (!driveUrl) return;
  try {
    const match = driveUrl.match(/[?&]id=([^&]+)/);
    if (!match) { telegramSendMessage('📷 ดูรูป: ' + driveUrl); return; }
    telegramSendPhoto(match[1]);
  } catch(e) { Logger.log('pushImageToAllSuperAdmins error: ' + e.message); telegramSendMessage('📷 ดูรูป: ' + driveUrl); }
}

// =================================================================================
// GOOGLE DRIVE
// =================================================================================
function savePhotoToDrive(base64, mimeType, tourId, checkpointId) {
  try {
    const iter = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
    const folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
    const ext = mimeType === 'image/png' ? '.png' : '.jpg';
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mimeType, `${tourId}_${checkpointId}_${Date.now()}${ext}`));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch(e) { Logger.log('savePhotoToDrive error: ' + e.message); return ''; }
}

// =================================================================================
// LINE API HELPERS
// =================================================================================
function replyMessage(replyToken, text) { replyMessageAdvanced(replyToken, [{ type:'text', text: String(text).substring(0,5000) }]); }

function replyMessageAdvanced(replyToken, messages) {
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ replyToken, messages }),
      headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN() }
    });
    if (res.getResponseCode() !== 200) Logger.log('LINE reply error: ' + res.getContentText());
  } catch(e) { Logger.log('replyMessageAdvanced error: ' + e.message); }
}

function pushMessage(userId, text) {
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ to: userId, messages: [{ type:'text', text: String(text).substring(0,5000) }] }),
      headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN() }
    });
    if (res.getResponseCode() !== 200) Logger.log('LINE push error: ' + res.getContentText());
  } catch(e) { Logger.log('pushMessage error: ' + e.message); }
}

function getUserProfile(userId) {
  try {
    const res = UrlFetchApp.fetch(`https://api.line.me/v2/bot/profile/${userId}`, {
      muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN() }
    });
    return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
  } catch(_) { return null; }
}

// =================================================================================
// TELEGRAM API
// =================================================================================
function telegramSendMessage(text) {
  try {
    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN()}/sendMessage`, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID(), text: String(text).substring(0,4096), parse_mode: 'HTML' }),
    });
    if (res.getResponseCode() !== 200) Logger.log('Telegram error: ' + res.getContentText());
  } catch(e) { Logger.log('telegramSendMessage error: ' + e.message); }
}

function telegramSendPhoto(fileId) {
  try {
    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN()}/sendPhoto`, {
      method: 'post', muteHttpExceptions: true,
      payload: { chat_id: TELEGRAM_CHAT_ID(), photo: DriveApp.getFileById(fileId).getBlob() },
    });
    if (res.getResponseCode() !== 200) Logger.log('Telegram sendPhoto error: ' + res.getContentText());
  } catch(e) {
    Logger.log('telegramSendPhoto error: ' + e.message);
    telegramSendMessage('📷 ดูรูปภาพ: https://drive.google.com/file/d/' + fileId + '/view');
  }
}

function telegramSendDocument(blob, caption) {
  try {
    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN()}/sendDocument`, {
      method: 'post', muteHttpExceptions: true,
      payload: { chat_id: TELEGRAM_CHAT_ID(), document: blob, caption: caption || '' },
    });
    if (res.getResponseCode() !== 200) Logger.log('Telegram sendDocument error: ' + res.getContentText());
    else Logger.log('Telegram sendDocument OK');
  } catch(e) { Logger.log('telegramSendDocument error: ' + e.message); }
}

// =================================================================================
// DASHBOARD API
// =================================================================================
function getTourData(p) {
  try {
    const startDate = p && p.startDate ? p.startDate : null;
    const endDate   = p && p.endDate   ? p.endDate   : null;
    const sheet = tourSheet(), data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'ok', total: 0, rows: [] };

    const rows = data.slice(1).map(r => ({
      tourId:         String(r[0]  || ''),
      checkpointId:   String(r[1]  || ''),
      checkpointName: String(r[2]  || ''),
      guardName:      String(r[3]  || ''),
      guardUserId:    String(r[4]  || ''),
      timestamp:      r[5]  ? formatISOFromSheet_(r[5])  : '',
      status:         String(r[6]  || ''),
      notes:          String(r[7]  || ''),
      photoUrl:       String(r[8]  || ''),
      tourStartTime:  r[9]  ? formatISOFromSheet_(r[9])  : '',
      gpsLat:         r[10] ? String(r[10]) : '',
      gpsLng:         r[11] ? String(r[11]) : '',
    }));

    let filtered = rows;
    if (startDate && endDate) {
      filtered = rows.filter(r => {
        if (!r.timestamp) return false;
        const datePart = r.timestamp.substring(0, 10);
        return datePart >= startDate && datePart <= endDate;
      });
    }

    filtered.sort((a, b) => {
      if (a.timestamp < b.timestamp) return 1;
      if (a.timestamp > b.timestamp) return -1;
      return 0;
    });

    return { status: 'ok', total: filtered.length, rows: filtered };
  } catch(e) {
    Logger.log('getTourData error: ' + e.message);
    return { status: 'error', message: e.message, rows: [] };
  }
}

// =================================================================================
// CHECKPOINT MANAGEMENT (Admin Panel)
// =================================================================================
function getCheckpointsAdmin() {
  try {
    const data = checkpointSheet().getDataRange().getValues();
    if (data.length <= 1) return { status: 'ok', checkpoints: [] };
    const cps = data.slice(1).filter(r => r[0]).map((r, i) => ({
      rowIndex: i + 2,   // 1-based row in sheet (header = row 1)
      id:       String(r[0]).trim(),
      name:     String(r[1]).trim(),
      location: String(r[2]).trim(),
      lat:      parseFloat(r[3]) || 0,
      lng:      parseFloat(r[4]) || 0,
      active:   String(r[5]).toUpperCase() !== 'FALSE',
      order:    Number(r[6]) || 99,
    })).sort((a, b) => a.order - b.order);
    return { status: 'ok', checkpoints: cps };
  } catch(e) {
    Logger.log('getCheckpointsAdmin error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function addCheckpoint(p) {
  try {
    const sheet = checkpointSheet();
    const id    = String(p.id || '').trim();
    const name  = String(p.name || '').trim();
    if (!id || !name) return { status: 'error', message: 'กรุณาระบุ CheckpointId และ CheckpointName' };
    // ตรวจสอบ ID ซ้ำ
    const data = sheet.getDataRange().getValues();
    if (data.slice(1).some(r => String(r[0]).trim() === id)) {
      return { status: 'error', message: `ID "${id}" มีอยู่แล้วในระบบ` };
    }
    const maxOrder = data.slice(1).reduce((m, r) => Math.max(m, Number(r[6]) || 0), 0);
    sheet.appendRow([
      id, name,
      String(p.location || '').trim(),
      parseFloat(p.lat) || 0,
      parseFloat(p.lng) || 0,
      'TRUE',
      maxOrder + 1,
    ]);
    Logger.log('addCheckpoint: ' + id + ' — ' + name);
    return { status: 'ok', message: `เพิ่มจุดตรวจ "${name}" เรียบร้อย` };
  } catch(e) {
    Logger.log('addCheckpoint error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function updateCheckpoint(p) {
  try {
    const sheet = checkpointSheet();
    const id    = String(p.id || '').trim();
    if (!id) return { status: 'error', message: 'กรุณาระบุ CheckpointId' };
    const data  = sheet.getDataRange().getValues();
    let rowIdx  = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) { rowIdx = i + 1; break; }
    }
    if (rowIdx < 0) return { status: 'error', message: `ไม่พบ ID "${id}"` };
    // อัปเดตเฉพาะ field ที่ส่งมา
    if (p.name     !== undefined) sheet.getRange(rowIdx, 2).setValue(String(p.name).trim());
    if (p.location !== undefined) sheet.getRange(rowIdx, 3).setValue(String(p.location).trim());
    if (p.lat      !== undefined) sheet.getRange(rowIdx, 4).setValue(parseFloat(p.lat) || 0);
    if (p.lng      !== undefined) sheet.getRange(rowIdx, 5).setValue(parseFloat(p.lng) || 0);
    if (p.active   !== undefined) sheet.getRange(rowIdx, 6).setValue(String(p.active) === 'false' ? 'FALSE' : 'TRUE');
    if (p.order    !== undefined) sheet.getRange(rowIdx, 7).setValue(Number(p.order) || 99);
    Logger.log('updateCheckpoint: ' + id);
    return { status: 'ok', message: `อัปเดตจุดตรวจ "${id}" เรียบร้อย` };
  } catch(e) {
    Logger.log('updateCheckpoint error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function deleteCheckpoint(checkpointId) {
  try {
    const id    = String(checkpointId || '').trim();
    if (!id) return { status: 'error', message: 'กรุณาระบุ CheckpointId' };
    const sheet = checkpointSheet();
    const data  = sheet.getDataRange().getValues();
    let rowIdx  = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) { rowIdx = i + 1; break; }
    }
    if (rowIdx < 0) return { status: 'error', message: `ไม่พบ ID "${id}"` };
    sheet.deleteRow(rowIdx);
    Logger.log('deleteCheckpoint: ' + id);
    return { status: 'ok', message: `ลบจุดตรวจ "${id}" เรียบร้อย` };
  } catch(e) {
    Logger.log('deleteCheckpoint error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function getActiveGuards() {
  try {
    const data = guardLogSheet().getDataRange().getValues(), now = new Date();
    const SESSION_MS = CONFIG.SESSION_HOURS * 3600 * 1000, guards = [];
    for (let i = data.length - 1; i > 0; i--) {
      const [userId, loginTime, userName, status, adminType] = data[i];
      if (status !== 'ACTIVE') continue;
      if (adminType === 'GUARD' && (now - new Date(loginTime)) > SESSION_MS) continue;
      if (guards.find(g => g.userId === userId)) continue;
      guards.push({ userId: String(userId||''), userName: String(userName||'—'), adminType: String(adminType||'GUARD'), loginTime: formatISOFromSheet_(loginTime) });
    }
    return { status: 'ok', guards };
  } catch(e) { Logger.log('getActiveGuards error: ' + e.message); return { status: 'error', message: e.message, guards: [] }; }
}

function getDashboardSummary() {
  try {
    const now = new Date();
    const todayTH = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const data = tourSheet().getDataRange().getValues();

    const getDateStr = cell => {
      if (!cell) return '';
      if (cell instanceof Date) return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      return String(cell).substring(0, 10);
    };

    const todayRows = data.slice(1).filter(r => getDateStr(r[5]) === todayTH);
    const sevenDaysAgo = Utilities.formatDate(new Date(now.getTime() - 7*24*3600*1000), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const weekRows = data.slice(1).filter(r => {
      const ds = getDateStr(r[5]);
      return ds >= sevenDaysAgo && ds <= todayTH;
    });

    return {
      status: 'ok',
      today: {
        date:     todayTH,
        checkins: todayRows.length,
        tours:    [...new Set(todayRows.map(r => r[0]))].length,
        abnormal: todayRows.filter(r => r[6] === 'ผิดปกติ').length,
        guards:   [...new Set(todayRows.map(r => r[3]))].length,
      },
      week: {
        checkins: weekRows.length,
        tours:    [...new Set(weekRows.map(r => r[0]))].length,
        abnormal: weekRows.filter(r => r[6] === 'ผิดปกติ').length,
      },
      total: {
        allCheckins: data.length - 1,
        allTours:    [...new Set(data.slice(1).map(r => r[0]))].length,
        allAbnormal: data.slice(1).filter(r => r[6] === 'ผิดปกติ').length,
      },
    };
  } catch(e) {
    Logger.log('getDashboardSummary error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function exportExcel(p) {
  try {
    const startDate = (p && p.startDate) ? p.startDate : Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const endDate   = (p && p.endDate)   ? p.endDate   : startDate;
    const rows      = (getTourData({ startDate, endDate }).rows) || [];
    if (rows.length === 0) return HtmlService.createHtmlOutput('<meta charset="UTF-8"><h3 style="font-family:sans-serif;color:#888">ไม่มีข้อมูลในช่วงวันที่ที่เลือก</h3>');
    const ss = SpreadsheetApp.create(`GuardTour_${startDate}_${endDate}`), sheet = ss.getActiveSheet();
    sheet.setName('รายงานตรวจรอบ');
    const headers = ['ลำดับ','วันที่','เวลา','รอบตรวจ (TourId)','รหัสจุด','ชื่อจุดตรวจ','ยาม','สถานะ','หมายเหตุ','GPS Lat','GPS Lng','ลิงก์รูปภาพ'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#1a1a2e').setFontColor('#f59e0b').setFontWeight('bold');
    sheet.setFrozenRows(1);
    const dataRows = rows.map((r,i) => { const dt = r.timestamp ? new Date(r.timestamp) : new Date(); return [i+1, Utilities.formatDate(dt,CONFIG.TIMEZONE,'dd/MM/yyyy'), Utilities.formatDate(dt,CONFIG.TIMEZONE,'HH:mm'), r.tourId||'', r.checkpointId||'', r.checkpointName||'', r.guardName||'', r.status||'', r.notes||'', r.gpsLat||'', r.gpsLng||'', r.photoUrl||'']; });
    if (dataRows.length > 0) sheet.getRange(2,1,dataRows.length,headers.length).setValues(dataRows);
    sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextContains('ผิดปกติ').setBackground('#ffcccc').setFontColor('#c00000').setRanges([sheet.getRange(2,8,Math.max(dataRows.length,1),1)]).build()]);
    for (let c=1;c<=headers.length;c++) sheet.autoResizeColumn(c);
    const ab = rows.filter(r=>r.status==='ผิดปกติ').length, tc = [...new Set(rows.map(r=>r.tourId))].length, gc = [...new Set(rows.map(r=>r.guardName))].length;
    sheet.getRange(dataRows.length+3,1).setValue(`สรุป: ${rows.length} รายการ | ${tc} รอบ | ${gc} ยาม | ผิดปกติ ${ab} จุด`).setFontWeight('bold').setFontColor('#f59e0b');
    const fileId = ss.getId(); DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const xlsxUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
    return HtmlService.createHtmlOutput(`<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:sans-serif;background:#0b0c10;color:#dde1ec;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}.box{text-align:center;padding:40px;background:#111218;border-radius:16px;border:1px solid rgba(245,166,35,.3);max-width:420px}h2{color:#f5a623;margin-bottom:12px}p{color:#5a5d72;font-size:14px;margin-bottom:20px}a{display:inline-block;background:#f5a623;color:#000;padding:12px 28px;border-radius:10px;text-decoration:none;font-weight:700;font-size:16px}a:hover{background:#ffd166}.note{font-size:12px;margin-top:16px;color:#3d3f52}</style></head><body><div class="box"><div style="font-size:48px;margin-bottom:12px">📊</div><h2>Guard Tour Report</h2><p>${startDate} ถึง ${endDate}<br>${rows.length} รายการ · ${tc} รอบ · ${gc} ยาม</p><a href="${xlsxUrl}" download="GuardTour_${startDate}_${endDate}.xlsx">⬇️ ดาวน์โหลด Excel</a><div class="note">กดปุ่มด้านบนหากไม่ดาวน์โหลดอัตโนมัติ</div></div><script>setTimeout(function(){window.location.href="${xlsxUrl}"},1500);</script></body></html>`);
  } catch(e) { Logger.log('exportExcel error: ' + e.message); return HtmlService.createHtmlOutput('<meta charset="UTF-8"><h3 style="color:red;font-family:sans-serif">Export Error</h3><pre>' + e.message + '</pre>'); }
}

// =================================================================================
// LATE ALERT
// =================================================================================
function checkLateAlert() {
  try {
    const now       = new Date();
    const todayDate = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const hh        = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH'));
    const mm        = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, 'mm'));
    const nowMin    = hh * 60 + mm;
    const yesterday = Utilities.formatDate(new Date(now.getTime() - 24*3600*1000), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const data = tourSheet().getDataRange().getValues();
    const recentRows = data.slice(1).filter(r => {
      if (!r[5]) return false;
      const ds = r[5] instanceof Date ? Utilities.formatDate(r[5], CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(r[5]).substring(0, 10);
      return ds === todayDate || ds === yesterday;
    });

    const alertLog = lateAlertLogSheet().getDataRange().getValues();
    const sentKeys = new Set(alertLog.slice(1).map(r => String(r[0])));
    const alerts = [];

    for (const sh of CONFIG.SHIFTS) {
      const [shH, shM] = sh.start.split(':').map(Number);
      const [eH,  eM]  = sh.end.split(':').map(Number);
      let shiftStartMin = shH * 60 + shM;
      let shiftEndMin   = eH  * 60 + eM;
      if (shiftEndMin <= shiftStartMin) shiftEndMin += 24 * 60;
      const graceMin = shiftStartMin + CONFIG.LATE_GRACE_MINUTES;
      let adjNow = nowMin;
      if (nowMin < shiftStartMin && shiftStartMin >= 12 * 60) adjNow += 24 * 60;
      if (adjNow < graceMin)         continue;
      if (adjNow > shiftEndMin + 120) continue;

      const hasCheckin = recentRows.some(r => {
        const cell = r[9] || r[5];
        if (!cell) return false;
        try {
          const rHH  = parseInt(cell instanceof Date ? Utilities.formatDate(cell, CONFIG.TIMEZONE, 'HH') : String(cell).substring(11, 13));
          const rMM  = parseInt(cell instanceof Date ? Utilities.formatDate(cell, CONFIG.TIMEZONE, 'mm') : String(cell).substring(14, 16));
          if (isNaN(rHH) || isNaN(rMM)) return false;
          const rMin = rHH * 60 + rMM;
          let adjR = rMin;
          if (adjR < shiftStartMin && shiftStartMin >= 12 * 60) adjR += 24 * 60;
          return adjR >= shiftStartMin && adjR < shiftEndMin;
        } catch(_) { return false; }
      });

      if (hasCheckin) continue;
      const alertKey = `${todayDate}_shift${sh.id}`;
      if (sentKeys.has(alertKey)) continue;
      alerts.push({ sh, alertKey });
    }

    if (alerts.length === 0) { Logger.log('checkLateAlert: ไม่พบผลัดที่ขาดการตรวจ'); return; }

    alerts.forEach(({ sh, alertKey }) => {
      const nowTimeStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');
      const msg =
        `🚨 แจ้งเตือน: ยามไม่ตรวจตามกำหนด!\n━━━━━━━━━━━━━━━━\n📅 วันที่: ${thaiDateShort(now)}\n🕐 เวลาปัจจุบัน: ${nowTimeStr}\n⏰ ผลัดที่ขาด: ${sh.label} (${sh.start}–${sh.end})\n━━━━━━━━━━━━━━━━\n⚠️ ไม่พบการ Check-in ในผลัดนี้\nกรุณาตรวจสอบและติดต่อยามโดยด่วน!`;
      pushToAllSuperAdmins(msg);
      Logger.log('checkLateAlert sent: ' + alertKey);
      lateAlertLogSheet().appendRow([alertKey, Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), sh.id, sh.label, msg]);
    });
  } catch(e) { Logger.log('checkLateAlert error: ' + e.message + '\n' + e.stack); }
}

function setupLateAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'checkLateAlert') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkLateAlert').timeBased().everyHours(1).create();
  Logger.log('✅ Late Alert Trigger ตั้งค่าแล้ว (ทุก 1 ชั่วโมง)');
}

function setupReportTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['autoSendDailyReport','autoSendMonthlyReport'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSendDailyReport').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('autoSendMonthlyReport').timeBased().onMonthDay(1).atHour(7).create();
  Logger.log('✅ Report Triggers ตั้งค่าแล้ว');
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('✅ ลบ Trigger ทั้งหมดแล้ว');
}

// =================================================================================
// STATS 30 DAYS
// =================================================================================
function getStats30Days() {
  try {
    const now  = new Date();
    const data = tourSheet().getDataRange().getValues();
    const days = {};
    const getDateStr = cell => {
      if (!cell) return '';
      if (cell instanceof Date) return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      return String(cell).substring(0, 10);
    };
    for (let i = 0; i < 30; i++) {
      const d = new Date(now); d.setDate(d.getDate() - i);
      const key = Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      days[key] = { date: key, checkins: 0, tours: new Set(), abnormal: 0, guards: new Set() };
    }
    data.slice(1).forEach(r => {
      const key = getDateStr(r[5]);
      if (!days[key]) return;
      days[key].checkins++;
      if (r[6] === 'ผิดปกติ') days[key].abnormal++;
      if (r[3]) days[key].guards.add(String(r[3]));
      if (r[0]) days[key].tours.add(String(r[0]));
    });
    const result = Object.values(days).map(d => ({ date: d.date, checkins: d.checkins, abnormal: d.abnormal, guards: d.guards.size, tours: d.tours.size })).sort((a, b) => a.date.localeCompare(b.date));
    return { status: 'ok', days: result };
  } catch(e) { Logger.log('getStats30Days error: ' + e.message); return { status: 'error', message: e.message, days: [] }; }
}

// =================================================================================
// GPS TRACKS
// =================================================================================
function getGpsTracks(p) {
  try {
    const data   = tourSheet().getDataRange().getValues();
    const tourId = p && p.tourId ? p.tourId : null;
    const date   = p && p.date   ? p.date   : null;
    let rows = data.slice(1).filter(r => r[0] && r[10] && r[11]);
    if (tourId)       { rows = rows.filter(r => String(r[0]) === tourId); }
    else if (date)    { rows = rows.filter(r => { try { return (r[5] instanceof Date ? Utilities.formatDate(r[5], CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(r[5]).substring(0,10)) === date; } catch(_) { return false; } }); }
    const tours = {};
    rows.forEach(r => {
      const tid = String(r[0]);
      if (!tours[tid]) tours[tid] = { tourId: tid, guardName: String(r[3]||''), startTime: formatISOFromSheet_(r[9]), points: [] };
      tours[tid].points.push({ checkpointId: String(r[1]||''), checkpointName: String(r[2]||''), lat: parseFloat(r[10])||0, lng: parseFloat(r[11])||0, timestamp: formatISOFromSheet_(r[5]), status: String(r[6]||'') });
    });
    return { status: 'ok', tours: Object.values(tours) };
  } catch(e) { Logger.log('getGpsTracks error: ' + e.message); return { status: 'error', message: e.message, tours: [] }; }
}

// =================================================================================
// DAILY / MONTHLY REPORT
// =================================================================================
function autoSendDailyReport() {
  try {
    const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
    const result = generateDailyReport(Utilities.formatDate(yesterday, CONFIG.TIMEZONE, 'yyyy-MM-dd'));
    Logger.log('autoSendDailyReport: ' + (result.status === 'ok' ? 'สำเร็จ' : result.message));
  } catch(e) { Logger.log('autoSendDailyReport error: ' + e.message); }
}

function autoSendMonthlyReport() {
  try {
    const lastMonth = new Date(); lastMonth.setDate(0);
    const result = generateMonthlyReport(String(lastMonth.getFullYear()), String(lastMonth.getMonth() + 1));
    Logger.log('autoSendMonthlyReport: ' + (result.status === 'ok' ? 'สำเร็จ' : result.message));
  } catch(e) { Logger.log('autoSendMonthlyReport error: ' + e.message); }
}

function generateDailyReport(dateStr) {
  try {
    if (!dateStr) dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const rows = (getTourData({ startDate: dateStr, endDate: dateStr }).rows) || [];
    const reportDate = new Date(dateStr + 'T12:00:00+07:00');
    const tours    = [...new Set(rows.map(r => r.tourId))];
    const abnormal = rows.filter(r => r.status === 'ผิดปกติ');
    const guards   = [...new Set(rows.map(r => r.guardName))];
    const shiftCounts = {};
    CONFIG.SHIFTS.forEach(s => { shiftCounts[s.id] = { label: s.label, start: s.start, end: s.end, count: 0 }; });
    rows.forEach(r => { if (!r.tourStartTime) return; const sh = detectShift_(new Date(r.tourStartTime)); if (sh) shiftCounts[sh.id].count++; });
    const cpMap = {};
    rows.forEach(r => { if (!cpMap[r.checkpointId]) cpMap[r.checkpointId] = { name: r.checkpointName, total: 0, abnormal: 0 }; cpMap[r.checkpointId].total++; if (r.status === 'ผิดปกติ') cpMap[r.checkpointId].abnormal++; });
    const dateDisplay = thaiDateFull(reportDate);
    let shiftSummary = '';
    CONFIG.SHIFTS.forEach(s => { const c = shiftCounts[s.id]; shiftSummary += `  ${c.label} (${c.start}-${c.end}): ${c.count > 0 ? c.count + ' จุด ✅' : '❌ ไม่มีการตรวจ'}\n`; });
    let cpSummary = Object.values(cpMap).map(c => `  📍 ${c.name}: ${c.total} ครั้ง${c.abnormal > 0 ? ` ⚠️ ผิดปกติ ${c.abnormal} จุด` : ''}`).join('\n');
    let abnSummary = '';
    if (abnormal.length > 0) { abnSummary = '\n━━━━━━━━━━━━━━━━\n⚠️ รายการผิดปกติ:\n'; abnSummary += abnormal.map(r => `  - ${r.checkpointName}: ${r.notes || '(ไม่ระบุ)'} [${r.guardName}]`).join('\n'); }
    const reportMsg =
      `📋 รายงานตรวจรอบประจำวัน\n━━━━━━━━━━━━━━━━\n📅 วันที่: ${dateDisplay}\n🏛️ หน่วย: ${CONFIG.UNIT_NAME}\n━━━━━━━━━━━━━━━━\n📊 สรุปผล:\n  🗂️ รอบตรวจ: ${tours.length} รอบ\n  📍 จุดตรวจทั้งหมด: ${rows.length} จุด\n  ⚠️ ผิดปกติ: ${abnormal.length} จุด\n  👤 ยาม: ${guards.join(', ') || 'ไม่มีข้อมูล'}\n━━━━━━━━━━━━━━━━\n🕐 สรุปรายผลัด:\n${shiftSummary}━━━━━━━━━━━━━━━━\n📍 สรุปรายจุด:\n${cpSummary || '  (ไม่มีข้อมูล)'}${abnSummary}\n━━━━━━━━━━━━━━━━\n${abnormal.length === 0 ? '✅ ปกติทุกจุด เหตุการณ์ทั่วไปปกติ' : '⚠️ มีรายการที่ต้องติดตาม'}`;
    pushToAllSuperAdmins(reportMsg);
    if (rows.length > 0) { try { sendDailyReportExcel_(rows, dateStr, dateDisplay); } catch(xlsErr) { Logger.log('sendDailyReportExcel_ error: ' + xlsErr.message); } }
    return { status: 'ok', date: dateStr, totalCheckins: rows.length, totalTours: tours.length, totalAbnormal: abnormal.length };
  } catch(e) { Logger.log('generateDailyReport error: ' + e.message + '\n' + e.stack); return { status: 'error', message: e.message }; }
}

function sendDailyReportExcel_(rows, dateStr, dateDisplay) {
  const ss    = SpreadsheetApp.create(`GuardTour_Daily_${dateStr}`);
  const sheet = ss.getActiveSheet();
  sheet.setName('รายงานตรวจรอบ');
  const headers = ['ลำดับ','เวลา','รอบตรวจ','จุดตรวจ','ยาม','สถานะ','หมายเหตุ','GPS Lat','GPS Lng'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#1a1a2e').setFontColor('#f59e0b').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(2,1,rows.length,headers.length).setValues(rows.map((r,i) => {
    const dt = r.timestamp ? new Date(r.timestamp) : new Date();
    return [i+1, Utilities.formatDate(dt,CONFIG.TIMEZONE,'HH:mm'), r.tourId||'', r.checkpointName||'', r.guardName||'', r.status||'', r.notes||'', r.gpsLat||'', r.gpsLng||''];
  }));
  for (let c=1;c<=headers.length;c++) sheet.autoResizeColumn(c);
  const fileId = ss.getId();
  DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  try {
    const xlsxRes  = UrlFetchApp.fetch(`https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    const xlsxBlob = xlsxRes.getBlob().setName(`GuardTour_${dateStr}.xlsx`);
    telegramSendDocument(xlsxBlob, `📊 รายงานตรวจรอบ ${dateDisplay}`);
  } catch(fetchErr) {
    Logger.log('sendDailyReportExcel_ fetch error: ' + fetchErr.message);
    telegramSendMessage(`📊 ดาวน์โหลดรายงาน Excel:\nhttps://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`);
  }
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(_) {}
}

function generateMonthlyReport(year, month) {
  try {
    if (!year || !month) { const d = new Date(); year = String(d.getFullYear()); month = String(d.getMonth() + 1); }
    const y = parseInt(year), m = parseInt(month);
    const startDate = `${y}-${String(m).padStart(2,'0')}-01`;
    const lastDay   = new Date(y, m, 0).getDate();
    const endDate   = `${y}-${String(m).padStart(2,'0')}-${lastDay}`;
    const rows      = (getTourData({ startDate, endDate }).rows) || [];
    const MONTH_NAMES = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthDisplay = `${MONTH_NAMES[m]} พ.ศ. ${y + 543}`;
    const tours    = [...new Set(rows.map(r => r.tourId))];
    const abnormal = rows.filter(r => r.status === 'ผิดปกติ');
    const guards   = [...new Set(rows.map(r => r.guardName))];
    const dailyCounts = {};
    for (let d = 1; d <= lastDay; d++) {
      const key = `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
      dailyCounts[key] = { day: d, checkins: 0, tours: new Set(), abnormal: 0 };
    }
    rows.forEach(r => { const key = r.timestamp ? r.timestamp.substring(0, 10) : ''; if (dailyCounts[key]) { dailyCounts[key].checkins++; if (r.tourId) dailyCounts[key].tours.add(r.tourId); if (r.status === 'ผิดปกติ') dailyCounts[key].abnormal++; } });
    const missingDays = Object.values(dailyCounts).filter(d => d.checkins === 0).map(d => d.day);
    const reportMsg =
      `📋 รายงานตรวจรอบประจำเดือน\n━━━━━━━━━━━━━━━━\n📅 เดือน: ${monthDisplay}\n🏛️ หน่วย: ${CONFIG.UNIT_NAME}\n━━━━━━━━━━━━━━━━\n📊 สรุปผล:\n  🗂️ รอบตรวจทั้งหมด: ${tours.length} รอบ\n  📍 Check-in ทั้งหมด: ${rows.length} ครั้ง\n  ⚠️ ผิดปกติ: ${abnormal.length} ครั้ง\n  👤 ยามที่ปฏิบัติหน้าที่: ${guards.length} นาย\n  📅 วันที่มีการตรวจ: ${lastDay - missingDays.length}/${lastDay} วัน\n` +
      (missingDays.length > 0 ? `  ❌ วันที่ไม่มีการตรวจ: ${missingDays.join(', ')}\n` : `  ✅ ครบทุกวัน\n`) +
      `━━━━━━━━━━━━━━━━\n👤 รายชื่อยาม: ${guards.join(', ') || 'ไม่มีข้อมูล'}\n━━━━━━━━━━━━━━━━\n${abnormal.length === 0 ? '✅ ไม่มีเหตุผิดปกติตลอดเดือน' : `⚠️ พบเหตุผิดปกติ ${abnormal.length} ครั้ง`}`;
    pushToAllSuperAdmins(reportMsg);
    if (rows.length > 0) { try { sendDailyReportExcel_(rows, `${y}-${String(m).padStart(2,'0')}`, `เดือน${monthDisplay}`); } catch(_) {} }
    return { status: 'ok', month: monthDisplay, totalCheckins: rows.length, totalTours: tours.length, totalAbnormal: abnormal.length, missingDays };
  } catch(e) { Logger.log('generateMonthlyReport error: ' + e.message + '\n' + e.stack); return { status: 'error', message: e.message }; }
}

// =================================================================================
// DEBUG
// =================================================================================
function getDebugInfo() {
  const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  const triggers = ScriptApp.getProjectTriggers().map(t => `${t.getHandlerFunction()} (${t.getTriggerSource()})`).join(', ') || 'ไม่มี';
  const props = PropertiesService.getScriptProperties().getProperties();
  const propStatus = ['SHEET_ID','CHANNEL_ACCESS_TOKEN','WEB_APP_URL','LIFF_ID','ADMIN_SECRET','SUPERADMIN_SECRET','TELEGRAM_BOT_TOKEN','TELEGRAM_CHAT_ID']
    .map(k => `${k}: ${props[k] ? '✅ ตั้งค่าแล้ว' : '❌ ยังไม่ได้ตั้งค่า'}`).join('<br>');
  return `<h2>🛡️ Guard Tour System v3.0</h2>
    <p><b>เวลาปัจจุบัน (TH):</b> ${now}</p>
    <p><b>Active Triggers:</b> ${triggers}</p>
    <h3>Script Properties Status:</h3><p>${propStatus}</p>
    <p><b>APIs:</b> get_tour_data, get_active_guards, get_dashboard_summary, export_excel, generate_daily_report, generate_monthly_report, get_stats_30days, get_gps_tracks</p>
    <pre style="background:#eee;padding:10px;font-size:12px">${Logger.getLog()}</pre>`;
}

// =================================================================================
// TEST FUNCTIONS
// =================================================================================
function testPushToSuperAdmin()    { pushToAllSuperAdmins('🧪 ทดสอบระบบแจ้งเตือน — Guard Tour System v3.0'); }
function testLateAlert()           { checkLateAlert(); Logger.log('testLateAlert done — ดู Telegram'); }
function testDailyReport()         { const r = generateDailyReport(); Logger.log(JSON.stringify(r)); }
function testMonthlyReport()       { const r = generateMonthlyReport(); Logger.log(JSON.stringify(r)); }
function testGetStats30Days()      { const r = getStats30Days(); Logger.log('30 days: ' + r.days.length + ' วัน'); }
function testGetGpsTracks()        { const r = getGpsTracks({}); Logger.log('GPS tours: ' + r.tours.length); }
function testGetTourData()         { const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); const r = getTourData({ startDate: today, endDate: today }); Logger.log('rows: ' + r.rows.length); }
function testGetActiveGuards()     { const r = getActiveGuards(); Logger.log('guards: ' + r.guards.length); }
function testGetDashboardSummary() { Logger.log(JSON.stringify(getDashboardSummary())); }
