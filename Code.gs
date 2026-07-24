/* =============================================================================
   Mini Farm Island 3D  —  Google Apps Script Backend (Code.gs)
   -----------------------------------------------------------------------------
   วิธีติดตั้ง:
   1. เปิด Google Sheet เป้าหมาย (ID: 1rKP0eVAdhZudGpeZHl9KF0Cf8FXnJcZ5wls-sWP21rA)
   2. เมนู Extensions > Apps Script  แล้ววางโค้ดนี้แทนของเดิมทั้งหมด
   3. กด Deploy > New deployment > เลือกชนิด "Web app"
        - Execute as:  Me (บัญชีคุณ)
        - Who has access:  Anyone            <-- สำคัญมาก ต้องเป็น Anyone
   4. คัดลอก URL /exec ที่ได้ ไปวางใน index.html (ตัวแปร APPS_SCRIPT_URL)
   5. ทุกครั้งที่แก้โค้ดนี้ ต้อง Deploy > Manage deployments > แก้ version เป็น New
      (ไม่งั้น URL เดิมจะยังรันโค้ดเวอร์ชันเก่า)
   ========================================================================== */

const SHEET_ID = '1rKP0eVAdhZudGpeZHl9KF0Cf8FXnJcZ5wls-sWP21rA';

// ---- Web App entry points --------------------------------------------------

function doPost(e) {
  try {
    const payload = parsePayload(e);
    const action = String(payload.action || '').trim();

    switch (action) {
      case 'register': handleAuth(payload, 'register'); break;
      case 'login':    handleAuth(payload, 'login');    break;
      case 'logout':   handleLogout(payload);           break;
      case 'saveGame': handleSaveGame(payload);         break;
      default:
        return jsonOut({ ok: false, error: 'unknown_action', action: action });
    }
    return jsonOut({ ok: true, action: action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// เปิด URL /exec ตรง ๆ ในเบราว์เซอร์เพื่อเช็คว่า deploy ทำงานอยู่
function doGet() {
  return jsonOut({ ok: true, status: 'Mini Farm 3D backend is running', time: new Date().toISOString() });
}

// ---- Payload parsing -------------------------------------------------------
// index.html ส่งมาแบบ x-www-form-urlencoded โดยมี field ชื่อ data = JSON ทั้งก้อน
// เผื่อกรณีอื่นด้วย: JSON body ตรง ๆ หรือ query string

function parsePayload(e) {
  if (e && e.parameter && e.parameter.data) {
    try { return JSON.parse(e.parameter.data); } catch (ignore) {}
  }
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (ignore) {}
  }
  return (e && e.parameter) ? e.parameter : {};
}

// ---- Handlers --------------------------------------------------------------

function handleAuth(p, mode) {
  const sheet = getSheet('Users', ['Username', 'Password', 'CreatedAt', 'LastLogin', 'LoginCount']);
  const row = findRowByValue(sheet, 1, p.username);
  const now = p.timestamp || new Date().toISOString();

  if (row === -1) {
    // ผู้ใช้ใหม่ (ทั้ง register และ login ที่ยังไม่มีบัญชี ให้สร้างให้เลย)
    sheet.appendRow([p.username || '', p.password || '', now, now, 1]);
  } else {
    // มีอยู่แล้ว: อัปเดตเวลาล็อกอินล่าสุด + นับจำนวนครั้ง
    const loginCount = Number(sheet.getRange(row, 5).getValue()) || 0;
    sheet.getRange(row, 4).setValue(now);
    sheet.getRange(row, 5).setValue(loginCount + 1);
  }
  logActivity(mode, p.username, now, '');
}

function handleLogout(p) {
  const now = p.timestamp || new Date().toISOString();
  logActivity('logout', p.username, now, 'playSeconds=' + (p.playDurationSeconds || 0));
}

function handleSaveGame(p) {
  const sheet = getSheet('SaveData',
    ['Username', 'UpdatedAt', 'Level', 'EXP', 'MaxEXP', 'Coins', 'Water',
     'UnlockedPlots', 'HouseLevel', 'HasDog', 'HasCat', 'PlaySeconds', 'RawJSON']);

  const d = p.saveData || {};
  const now = p.timestamp || new Date().toISOString();
  const rowValues = [
    p.username || '',
    now,
    d.level || 0,
    d.exp || 0,
    d.maxExp || 0,
    d.coins || 0,
    d.water || 0,
    d.unlockedMainPlots || 0,
    d.houseLevel || 0,
    d.hasDog ? 'YES' : 'NO',
    d.hasCat ? 'YES' : 'NO',
    p.playDurationSeconds || 0,
    JSON.stringify(d)
  ];

  // Upsert: 1 แถวต่อ 1 ผู้ใช้ (เซฟทับของเดิม ไม่ให้ข้อมูลบวมหลายพันแถว)
  const row = findRowByValue(sheet, 1, p.username);
  if (row === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
  }
  logActivity('saveGame', p.username, now, 'coins=' + (d.coins || 0) + ', level=' + (d.level || 0));
}

// ---- Helpers ---------------------------------------------------------------

function logActivity(action, username, timestamp, detail) {
  const sheet = getSheet('ActivityLogs', ['Timestamp', 'Action', 'Username', 'Detail']);
  sheet.appendRow([timestamp || new Date().toISOString(), action, username || '', detail || '']);
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// หาแถวจากค่าในคอลัมน์ (คืน index แถวจริงเริ่มที่ 2, ไม่พบคืน -1)
function findRowByValue(sheet, col, value) {
  if (!value) return -1;
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const values = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return -1;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
