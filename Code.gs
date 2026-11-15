/**
 * บันทึกบิล — ตัวรับข้อมูลซิงก์ (Google Apps Script)
 *
 * วิธีติดตั้ง
 * 1) สร้าง Google Sheet ใหม่ > เมนู ส่วนขยาย (Extensions) > Apps Script
 * 2) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้แทน
 * 3) เปลี่ยนค่า SECRET ด้านล่างเป็นรหัสลับของคุณเอง (ยาว เดายาก) แล้วกดบันทึก
 * 4) Deploy > New deployment > เลือกประเภท Web app
 *      Execute as: Me   |   Who has access: Anyone
 *    กด Deploy แล้วอนุญาตสิทธิ์ (Authorize access)
 * 5) คัดลอก Web app URL (ลงท้ายด้วย /exec) ไปวางในแอป พร้อมรหัสลับเดียวกัน
 *
 * หมายเหตุ: ถ้าแก้โค้ดนี้ภายหลัง ต้อง Deploy > Manage deployments > Edit > Version: New version
 */

const SECRET = 'CHANGE-ME';
const SHEET_NAME = 'รายจ่าย';
const HEADERS = ['ID', 'วันที่', 'เวลา', 'ร้าน', 'จำนวนเงิน', 'หมวดหมู่', 'ช่องทางจ่าย', 'แท็ก', 'รายละเอียด', 'รูปบิล', 'สร้างเมื่อ', 'แก้ไขล่าสุด'];

function doGet() {
  return out_({ ok: true, app: 'billbook', message: 'พร้อมรับข้อมูล' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad_json' });
  }
  if (!SECRET || SECRET === 'CHANGE-ME') {
    return out_({ ok: false, error: 'ยังไม่ได้ตั้งค่า SECRET ใน Apps Script' });
  }
  if (body.secret !== SECRET) return out_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return out_({ ok: false, error: 'busy' });
  try {
    const sh = sheet_();
    if (body.action === 'ping') {
      return out_({ ok: true, rows: Math.max(sh.getLastRow() - 1, 0) });
    }
    if (body.action === 'sync') {
      const res = applyOps_(sh, body.ops || []);
      res.ok = true;
      return out_(res);
    }
    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('#,##0.00');
  }
  return sh;
}

function applyOps_(sh, ops) {
  const last = sh.getLastRow();
  const index = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
      const id = String(r[0]);
      if (id) index[id] = i + 2;
    });
  }
  const updates = [];
  const delRows = [];
  const appendMap = {};
  const appends = [];

  ops.forEach(function (op) {
    if (!op) return;
    if (op.op === 'delete') {
      const id = String(op.id);
      if (index[id]) { delRows.push(index[id]); delete index[id]; }
      if (appendMap[id] !== undefined) { appends[appendMap[id]] = null; delete appendMap[id]; }
    } else if (op.op === 'upsert' && op.item && op.item.id) {
      const id = String(op.item.id);
      const row = row_(op.item);
      if (index[id]) updates.push([index[id], row]);
      else if (appendMap[id] !== undefined) appends[appendMap[id]] = row;
      else { appendMap[id] = appends.length; appends.push(row); }
    }
  });

  updates.forEach(function (u) {
    sh.getRange(u[0], 1, 1, u[1].length).setValues([u[1]]);
  });

  delRows.sort(function (a, b) { return b - a; }).forEach(function (r) {
    // ชีตห้ามลบแถวที่ไม่ได้ freeze จนหมด — กรณีเหลือแถวเดียวให้ล้างค่าแทน
    if (sh.getMaxRows() <= sh.getFrozenRows() + 1) {
      sh.getRange(r, 1, 1, HEADERS.length).clearContent();
    } else {
      sh.deleteRow(r);
    }
  });

  const rows = appends.filter(Boolean);
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return { updated: updates.length, deleted: delRows.length, added: rows.length };
}

function row_(it) {
  return [
    String(it.id),
    it.date || '',
    it.time || '',
    safe_(it.shop),
    Number(it.amount) || 0,
    safe_(it.category),
    safe_(it.payment),
    safe_(it.tags),
    safe_(it.note),
    it.hasImage ? 'มี (เก็บในมือถือ)' : '',
    it.createdAt || '',
    it.updatedAt || ''
  ];
}

// กันข้อความที่ขึ้นต้นด้วย = + - @ ถูกตีความเป็นสูตร
function safe_(v) {
  const s = v == null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
