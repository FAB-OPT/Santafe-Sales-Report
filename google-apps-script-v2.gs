/****************************************************
 * Santa Fe Sales Submit – Google Apps Script v2.0
 *
 * Version: 2.0.0 (21.4.2026)
 * รวม: Web App API + History + Edit + Telegram Notify
 *
 * การ Deploy:
 * 1. คัดลอกไฟล์นี้ทั้งหมดไปใส่ใน Google Apps Script Editor
 * 2. ตั้งค่า Script Properties:
 *    - TELEGRAM_NOTIFY_URL = https://script.google.com/macros/s/XXXXX/exec
 * 3. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. คัดลอก URL ไปใส่ใน index v2.0.html (WEB_APP_URL)
 ****************************************************/

const API_VERSION        = "2.0.0 (21.4.2026)";
const API_SPREADSHEET_ID = "1GBrc4S2lSu40bMOelAHvEzRGf7EIBhFGOKSISa-EZRU";
const API_DATA_SHEET_NAME = "DATA";
const API_TIMEZONE       = "Asia/Bangkok";
const API_LOCK_WAIT_MS   = 120000;

const PROP_TG_NOTIFY_URL = "TELEGRAM_NOTIFY_URL";

const API_TYPE_HEADER        = "type";
const API_TYPE_SANTA_FE      = "Santa fe";
const API_TYPE_SANTA_FE_EASY = "Santa fe Easy";
const API_EASY_BRANCH_CODES  = { "5504": true, "5505": true, "5508": true, "5509": true };

/* =========================
   WebApp Entrypoints
========================= */
function doGet(e) {
  const p  = (e && e.parameter) ? e.parameter : {};
  const cb = p.callback ? String(p.callback) : "";
  const out = apiDispatch_(p);
  return apiJsonOut_(out, cb);
}

function doPost(e) {
  const params    = apiParseBody_(e);
  const transport = String(params._transport || "").trim().toLowerCase();
  const reqId     = String(params._reqId || "").trim();

  delete params._transport;
  delete params._reqId;

  const out = apiDispatch_(params);

  if (transport === "iframe") {
    return apiHtmlBridgeOut_(out, reqId);
  }
  return apiJsonOut_(out, "");
}

/* =========================
   ★ apiDispatch_ v2.0
   เพิ่ม mode: history, edit
========================= */
function apiDispatch_(p) {
  const mode = String((p && p.mode) || "").trim();

  if (mode === "ping") {
    return { ok: true, version: API_VERSION, message: "PING OK" };
  }

  if (mode === "checkdup") {
    const dm = String(p.district_manager || "");
    const br = String(p.branch || "");
    const dt = (p.submit_date !== undefined) ? p.submit_date : "";
    const sl = String(p.submit_time_slot || "");

    const sh  = apiGetSheet_();
    const hit = apiFindDupRowAny_(sh, dm, br, dt, sl);
    const dupRow = hit ? hit.row : null;
    const key    = hit ? hit.key : apiMakeKey_(dm, br, dt, sl);
    const dupObj = dupRow ? apiBuildDuplicateObj_(sh, dupRow, key) : null;

    return { ok: true, version: API_VERSION, duplicated: !!dupRow, duplicate: dupObj };
  }

  if (mode === "submit") {
    return apiHandleSubmit_(p);
  }

  // ★ v2.0: History
  if (mode === "history") {
    return apiHandleHistory_(p);
  }

  // ★ v2.0: Edit
  if (mode === "edit") {
    return apiHandleEdit_(p);
  }

  return { ok: true, version: API_VERSION, message: "OK" };
}

/* =========================
   Core submit handler
========================= */
function apiHandleSubmit_(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(API_LOCK_WAIT_MS);
  } catch (_) {
    return {
      ok: false, version: API_VERSION,
      code: "BUSY_PEAK",
      error: "busy: cannot acquire lock, please retry",
      userMessage: "ระบบกำลังมีการใช้งานสูง กรุณารอแล้วลองใหม่อีกครั้ง"
    };
  }

  let resp = null;
  let notifyPayload = null;

  try {
    const submitter = String(p.submitter_name || "").trim();
    const dm        = String(p.district_manager || "").trim();
    const br        = String(p.branch || "").trim();
    const date      = (p.submit_date !== undefined) ? p.submit_date : "";
    const slot      = String(p.submit_time_slot || "").trim();

    if (!submitter || !dm || !br || !date || !slot) {
      resp = { ok: false, version: API_VERSION, error: "missing required field" };
      return resp;
    }

    const sh         = apiGetSheet_();
    const canonDate  = apiCanonDate_(date);
    const canonSlot  = apiCanonSlot_(slot);

    // Rule #0: ก่อนส่ง 16.00 ต้องมีสิ้นวันของวันก่อน
    if (canonSlot === "16.00") {
      const prevDate = apiPrevDate_(canonDate);
      if (prevDate) {
        const hitPrevEod = apiFindDupRowAny_(sh, dm, br, prevDate, "สิ้นวัน");
        if (!hitPrevEod) {
          resp = {
            ok: false, version: API_VERSION,
            code: "MUST_SEND_PREV_EOD_FIRST",
            userMessage:
              `ยังไม่ได้ส่งรอบ "สิ้นวัน" ของวันที่ ${prevDate} ` +
              `จึงยังส่งรอบ 16.00 ของวันที่ ${canonDate} ไม่ได้ ` +
              `กรุณาย้อนกลับไปส่งสิ้นวันก่อน`,
            requiredDate: prevDate,
            requiredSlot: "สิ้นวัน",
            requiredKey: apiMakeKey_(dm, br, prevDate, "สิ้นวัน")
          };
          return resp;
        }
      }
    }

    // Rule #1: สิ้นวันต้องมี 16.00 ก่อน
    if (canonSlot === "สิ้นวัน") {
      const hit1600 = apiFindDupRowAny_(sh, dm, br, canonDate, "16.00");
      if (!hit1600) {
        resp = {
          ok: false, version: API_VERSION,
          code: "MUST_SEND_1600_FIRST",
          userMessage: "ต้องส่งข้อมูลรอบเวลา 16.00 ก่อนจึงจะส่งรอบ สิ้นวัน ได้",
          requiredKey: apiMakeKey_(dm, br, canonDate, "16.00")
        };
        return resp;
      }
    }

    // Rule #2: Labour hour < Labour baht (เฉพาะสิ้นวัน)
    let labourHour = apiNum_(p.labour_hour);
    let labourBaht = apiNum_(p.labour_baht);

    if (canonSlot === "16.00") {
      labourHour = 0;
      labourBaht = 0;
      p.labour_hour = 0;
      p.labour_baht = 0;
    } else {
      if (!(labourHour < labourBaht)) {
        resp = {
          ok: false, version: API_VERSION,
          code: "LABOUR_HOUR_GE_BAHT",
          userMessage: "เงื่อนไขไม่ผ่าน: Labour (hour) ต้องน้อยกว่า Labour (Baht) เสมอ",
          labour_hour: labourHour, labour_baht: labourBaht
        };
        return resp;
      }
    }

    // Dup check
    const hitDup = apiFindDupRowAny_(sh, dm, br, canonDate, canonSlot);
    if (hitDup) {
      resp = {
        ok: false, version: API_VERSION,
        duplicated: true,
        userMessage: "ข้อมูลซ้ำ: มีการส่งข้อมูลของ DM/สาขา/วันที่/ช่วงเวลา นี้แล้ว",
        duplicate: apiBuildDuplicateObj_(sh, hitDup.row, hitDup.key)
      };
      return resp;
    }

    // Append
    const tsText     = Utilities.formatDate(new Date(), API_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    const keyStd     = apiMakeKey_(dm, br, canonDate, canonSlot);
    const branchType = apiResolveType_(br);

    const rowValues = [
      tsText, submitter, apiNorm_(dm), apiNorm_(br), canonDate, canonSlot,
      apiNum_(p.plan_sale), apiNum_(p.actual_sale),
      apiNum_(p.sale_dine_in), apiNum_(p.sale_take_away),
      apiNum_(p.sale_grab), apiNum_(p.sale_lineman), apiNum_(p.sale_shopeefood),
      apiInt_(p.total_trans), apiInt_(p.trans_dine_in), apiInt_(p.trans_take_away),
      apiInt_(p.trans_grab), apiInt_(p.trans_lineman), apiInt_(p.trans_shopeefood),
      apiInt_(p.customer), labourHour, labourBaht, keyStd, branchType
    ];

    sh.appendRow(rowValues);
    SpreadsheetApp.flush();

    notifyPayload = {
      submitter_name: submitter,
      district_manager: apiNorm_(dm),
      branch: apiNorm_(br),
      branch_code: apiBranchCode_(br),
      submit_date: canonDate,
      submit_time_slot: canonSlot,
      plan_sale: apiNum_(p.plan_sale),
      actual_sale: apiNum_(p.actual_sale),
      total_trans: apiInt_(p.total_trans),
      customer: apiInt_(p.customer),
      labour_hour: labourHour,
      labour_baht: labourBaht,
      dup_key: keyStd,
      type: branchType
    };

    resp = {
      ok: true, version: API_VERSION,
      sheetWritten: true,
      telegramSent: false, telegramError: null,
      dupKey: keyStd, type: branchType
    };
    return resp;

  } catch (err) {
    resp = { ok: false, version: API_VERSION, error: String(err && err.stack ? err.stack : err) };
    return resp;

  } finally {
    try { lock.releaseLock(); } catch (_) {}
    if (resp && resp.ok === true && resp.sheetWritten === true) {
      try {
        const notifyUrl = PropertiesService.getScriptProperties().getProperty(PROP_TG_NOTIFY_URL);
        if (!notifyUrl) {
          resp.telegramSent = false;
          resp.telegramError = "Missing Script Property: " + PROP_TG_NOTIFY_URL;
        } else {
          const res  = UrlFetchApp.fetch(notifyUrl, {
            method: "post", contentType: "application/json",
            payload: JSON.stringify(notifyPayload || {}), muteHttpExceptions: true
          });
          const code = res.getResponseCode();
          if (code < 200 || code >= 300) throw new Error("Notify HTTP " + code + ": " + res.getContentText());
          resp.telegramSent  = true;
          resp.telegramError = null;
        }
      } catch (e) {
        resp.telegramSent  = false;
        resp.telegramError = String(e && e.stack ? e.stack : e);
      }
    }
  }
}

/* =========================
   ★ v2.0: History Handler
   ดึงข้อมูลย้อนหลัง N วันของสาขา
========================= */
function apiHandleHistory_(p) {
  try {
    const branchCode = String(p.branch_code || "").trim();
    const days       = parseInt(p.days) || 30;

    if (!branchCode) {
      return { ok: false, version: API_VERSION, error: "missing branch_code" };
    }

    const sh   = apiGetSheet_();
    const last = sh.getLastRow();
    if (last < 2) {
      return { ok: true, version: API_VERSION, rows: [], total: 0 };
    }

    const now      = new Date();
    const cutoff   = new Date(now.getTime() - (days * 24 * 60 * 60 * 1000));
    const cutoffStr = Utilities.formatDate(cutoff, API_TIMEZONE, "yyyy-MM-dd");

    const numRows  = last - 1;
    const numCols  = sh.getLastColumn();
    const allData  = sh.getRange(2, 1, numRows, numCols).getValues();
    const headers  = sh.getRange(1, 1, 1, numCols).getValues()[0]
      .map(v => String(v || "").trim());

    const results = [];

    for (let i = 0; i < allData.length; i++) {
      const row      = allData[i];
      const brFull   = String(row[3] || "").trim(); // col D = branch
      const dateVal  = row[4];                       // col E = submit_date

      const rowCode = apiBranchCode_(brFull);
      if (rowCode !== branchCode) continue;

      let dateStr;
      if (Object.prototype.toString.call(dateVal) === "[object Date]" && !isNaN(dateVal.getTime())) {
        dateStr = Utilities.formatDate(dateVal, API_TIMEZONE, "yyyy-MM-dd");
      } else {
        dateStr = apiCanonDate_(dateVal);
      }
      if (!dateStr || dateStr < cutoffStr) continue;

      const entry = { _row: i + 2 };
      for (let c = 0; c < headers.length; c++) {
        const h = headers[c];
        if (!h) continue;
        let val = row[c];
        if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
          val = (h === "submit_date")
            ? Utilities.formatDate(val, API_TIMEZONE, "yyyy-MM-dd")
            : Utilities.formatDate(val, API_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
        }
        entry[h] = val;
      }
      results.push(entry);
    }

    results.sort((a, b) => {
      const da = String(a.submit_date || "");
      const db = String(b.submit_date || "");
      if (da !== db) return db.localeCompare(da);
      return String(b.submit_time_slot || "").localeCompare(String(a.submit_time_slot || ""));
    });

    return {
      ok: true, version: API_VERSION,
      rows: results, total: results.length,
      branch_code: branchCode, cutoff_date: cutoffStr
    };

  } catch (err) {
    return { ok: false, version: API_VERSION, error: String(err) };
  }
}

/* =========================
   ★ v2.0: Edit Handler
   แก้ไขแถวเดิม + เก็บ backup + Telegram
========================= */
function apiHandleEdit_(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(API_LOCK_WAIT_MS);
  } catch (_) {
    return {
      ok: false, version: API_VERSION,
      code: "BUSY_PEAK",
      userMessage: "ระบบกำลังมีการใช้งานสูง กรุณารอแล้วลองใหม่"
    };
  }

  let resp          = null;
  let notifyPayload = null;

  try {
    const targetRow  = parseInt(p._row);
    const branchCode = String(p.branch_code || "").trim();

    if (!targetRow || targetRow < 2) {
      resp = { ok: false, version: API_VERSION, error: "missing or invalid _row" };
      return resp;
    }
    if (!branchCode) {
      resp = { ok: false, version: API_VERSION, error: "missing branch_code for auth" };
      return resp;
    }

    const sh      = apiGetSheet_();
    const lastRow = sh.getLastRow();

    if (targetRow > lastRow) {
      resp = { ok: false, version: API_VERSION, error: "row not found" };
      return resp;
    }

    // ตรวจว่าแถวนี้เป็นของสาขาที่ล็อกอิน
    const existingBranch = String(sh.getRange(targetRow, 4).getValue() || "").trim();
    const existingCode   = apiBranchCode_(existingBranch);
    if (existingCode !== branchCode) {
      resp = {
        ok: false, version: API_VERSION,
        code: "UNAUTHORIZED",
        userMessage: "ไม่สามารถแก้ไขข้อมูลของสาขาอื่นได้"
      };
      return resp;
    }

    // อ่านข้อมูลเดิม (backup)
    const numCols   = sh.getLastColumn();
    const oldValues = sh.getRange(targetRow, 1, 1, numCols).getValues()[0];
    const headers   = sh.getRange(1, 1, 1, numCols).getValues()[0]
      .map(v => String(v || "").trim());

    const backupObj = {};
    for (let c = 0; c < headers.length; c++) {
      if (headers[c]) {
        let val = oldValues[c];
        if (Object.prototype.toString.call(val) === "[object Date]") {
          val = Utilities.formatDate(val, API_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
        }
        backupObj[headers[c]] = val;
      }
    }

    // Validate required fields
    const submitter  = String(p.submitter_name || "").trim();
    const dm         = String(p.district_manager || "").trim();
    const br         = String(p.branch || "").trim();
    const date       = String(p.submit_date || "").trim();
    const slot       = String(p.submit_time_slot || "").trim();

    if (!submitter || !dm || !br || !date || !slot) {
      resp = { ok: false, version: API_VERSION, error: "missing required field" };
      return resp;
    }

    const canonDate = apiCanonDate_(date);
    const canonSlot = apiCanonSlot_(slot);

    let labourHour = apiNum_(p.labour_hour);
    let labourBaht = apiNum_(p.labour_baht);

    if (canonSlot === "16.00") {
      labourHour = 0;
      labourBaht = 0;
    } else {
      if (!(labourHour < labourBaht)) {
        resp = {
          ok: false, version: API_VERSION,
          code: "LABOUR_HOUR_GE_BAHT",
          userMessage: "Labour (hour) ต้องน้อยกว่า Labour (Baht)"
        };
        return resp;
      }
    }

    // เขียนข้อมูลใหม่
    const tsText     = Utilities.formatDate(new Date(), API_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    const branchType = apiResolveType_(br);

    const newRowValues = [
      tsText, submitter, apiNorm_(dm), apiNorm_(br), canonDate, canonSlot,
      apiNum_(p.plan_sale), apiNum_(p.actual_sale),
      apiNum_(p.sale_dine_in), apiNum_(p.sale_take_away),
      apiNum_(p.sale_grab), apiNum_(p.sale_lineman), apiNum_(p.sale_shopeefood),
      apiInt_(p.total_trans), apiInt_(p.trans_dine_in), apiInt_(p.trans_take_away),
      apiInt_(p.trans_grab), apiInt_(p.trans_lineman), apiInt_(p.trans_shopeefood),
      apiInt_(p.customer), labourHour, labourBaht,
      apiMakeKey_(dm, br, canonDate, canonSlot), branchType
    ];

    const writeLen = Math.min(newRowValues.length, numCols);
    sh.getRange(targetRow, 1, 1, writeLen).setValues([newRowValues.slice(0, writeLen)]);

    // Edit metadata columns
    apiEnsureEditColumns_(sh);
    const editCountCol     = apiGetColByHeader_(sh, "edit_count");
    const editTimestampCol = apiGetColByHeader_(sh, "edit_timestamp");
    const originalDataCol  = apiGetColByHeader_(sh, "original_data");
    const lastEditorCol    = apiGetColByHeader_(sh, "last_editor");

    let editCount = 0;
    try { editCount = parseInt(sh.getRange(targetRow, editCountCol).getValue()) || 0; } catch (_) {}

    if (editCount === 0 && originalDataCol) {
      sh.getRange(targetRow, originalDataCol).setValue(JSON.stringify(backupObj));
    }

    editCount++;
    if (editCountCol)     sh.getRange(targetRow, editCountCol).setValue(editCount);
    if (editTimestampCol) sh.getRange(targetRow, editTimestampCol).setValue(tsText);
    if (lastEditorCol)    sh.getRange(targetRow, lastEditorCol).setValue(submitter);

    SpreadsheetApp.flush();

    notifyPayload = {
      submitter_name: submitter,
      district_manager: apiNorm_(dm),
      branch: apiNorm_(br),
      branch_code: branchCode,
      submit_date: canonDate,
      submit_time_slot: canonSlot,
      plan_sale: apiNum_(p.plan_sale),
      actual_sale: apiNum_(p.actual_sale),
      total_trans: apiInt_(p.total_trans),
      customer: apiInt_(p.customer),
      labour_hour: labourHour,
      labour_baht: labourBaht,
      type: branchType,
      _edited: true,
      _edit_count: editCount
    };

    resp = {
      ok: true, version: API_VERSION,
      edited: true, editedRow: targetRow, editCount: editCount,
      telegramSent: false, telegramError: null
    };
    return resp;

  } catch (err) {
    resp = { ok: false, version: API_VERSION, error: String(err && err.stack ? err.stack : err) };
    return resp;

  } finally {
    try { lock.releaseLock(); } catch (_) {}
    if (resp && resp.ok === true && resp.edited === true) {
      try {
        const notifyUrl = PropertiesService.getScriptProperties().getProperty(PROP_TG_NOTIFY_URL);
        if (notifyUrl && notifyPayload) {
          const res  = UrlFetchApp.fetch(notifyUrl, {
            method: "post", contentType: "application/json",
            payload: JSON.stringify(notifyPayload), muteHttpExceptions: true
          });
          const code = res.getResponseCode();
          if (code >= 200 && code < 300) {
            resp.telegramSent = true;
          } else {
            resp.telegramError = "HTTP " + code;
          }
        }
      } catch (e) {
        resp.telegramSent  = false;
        resp.telegramError = String(e);
      }
    }
  }
}

/* =========================
   Edit column helpers
========================= */
function apiEnsureEditColumns_(sh) {
  const editCols = ["edit_count", "edit_timestamp", "original_data", "last_editor"];
  const lastCol  = sh.getLastColumn();
  const hdr      = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  editCols.forEach(col => {
    if (hdr.indexOf(col) < 0) {
      const newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(col);
    }
  });
}

function apiGetColByHeader_(sh, headerName) {
  const lastCol = sh.getLastColumn();
  const hdr     = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const idx     = hdr.indexOf(headerName);
  return idx >= 0 ? idx + 1 : null;
}

/* =========================
   Sheet helpers
========================= */
function apiGetSheet_() {
  const ss = SpreadsheetApp.openById(API_SPREADSHEET_ID);
  let sh   = ss.getSheetByName(API_DATA_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(API_DATA_SHEET_NAME);
  apiEnsureHeaderAndDupCol_(sh);
  return sh;
}

function apiEnsureHeaderAndDupCol_(sh) {
  const headers = [
    "timestamp","submitter_name","district_manager","branch","submit_date","submit_time_slot",
    "plan_sale","actual_sale","sale_dine_in","sale_take_away","sale_grab","sale_lineman","sale_shopeefood",
    "total_trans","trans_dine_in","trans_take_away","trans_grab","trans_lineman","trans_shopeefood",
    "customer","labour_hour","labour_baht","__dup_key","type"
  ];

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), headers.length);

  if (lastRow === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const cur     = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  let changed   = false;
  for (let i = 0; i < headers.length; i++) {
    if (cur[i] !== headers[i]) { cur[i] = headers[i]; changed = true; }
  }
  if (changed) {
    sh.getRange(1, 1, 1, Math.max(cur.length, headers.length)).setValues([
      cur.slice(0, Math.max(cur.length, headers.length))
    ]);
  }
  apiDupCol_(sh);
  apiTypeCol_(sh);
}

function apiDupCol_(sh) {
  const lastCol = sh.getLastColumn();
  const hdr     = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const idx     = hdr.indexOf("__dup_key");
  if (idx >= 0) return idx + 1;
  const newCol  = lastCol + 1;
  sh.getRange(1, newCol).setValue("__dup_key");
  return newCol;
}

function apiTypeCol_(sh) {
  const lastCol = sh.getLastColumn();
  const hdr     = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const idx     = hdr.indexOf(API_TYPE_HEADER);
  if (idx >= 0) return idx + 1;
  const newCol  = Math.max(lastCol + 1, apiDupCol_(sh) + 1);
  sh.getRange(1, newCol).setValue(API_TYPE_HEADER);
  return newCol;
}

/* =========================
   Duplicate finders
========================= */
function apiFindDupRow_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const col   = apiDupCol_(sh);
  const found = sh.getRange(2, col, last - 1, 1).createTextFinder(String(key)).matchEntireCell(true).findNext();
  return found ? found.getRow() : null;
}

function apiFindDupRowAny_(sh, dm, br, dt, sl) {
  const canonDate = apiCanonDate_(dt);
  const canonSlot = apiCanonSlot_(sl);
  const keys = [apiMakeKey_(dm, br, canonDate, canonSlot)];
  if (canonSlot === "16.00") keys.push(apiMakeKeyRawSlot_(dm, br, canonDate, "16"));
  if (String(sl || "").trim() === "16") keys.push(apiMakeKey_(dm, br, canonDate, "16.00"));
  for (let i = 0; i < keys.length; i++) {
    const row = apiFindDupRow_(sh, keys[i]);
    if (row) return { row: row, key: keys[i] };
  }
  return null;
}

/* =========================
   Admin helpers (run once)
========================= */
function adminRebuildDupKeysAll_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(API_LOCK_WAIT_MS);
  try {
    const sh   = apiGetSheet_();
    const last = sh.getLastRow();
    if (last < 2) return;
    const dupCol = apiDupCol_(sh);
    const dmVals = sh.getRange(2, 3, last - 1, 1).getValues();
    const brVals = sh.getRange(2, 4, last - 1, 1).getValues();
    const dtVals = sh.getRange(2, 5, last - 1, 1).getValues();
    const slVals = sh.getRange(2, 6, last - 1, 1).getValues();
    const out    = [];
    for (let i = 0; i < last - 1; i++) {
      out.push([apiMakeKey_(dmVals[i][0], brVals[i][0], dtVals[i][0], slVals[i][0])]);
    }
    sh.getRange(2, dupCol, last - 1, 1).setValues(out);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function adminBackfillTypeAll() {
  const lock = LockService.getScriptLock();
  lock.waitLock(API_LOCK_WAIT_MS);
  try {
    const sh   = apiGetSheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, version: API_VERSION, updated: 0 };
    const typeCol = apiTypeCol_(sh);
    const brVals  = sh.getRange(2, 4, last - 1, 1).getValues();
    const out     = brVals.map(r => [apiResolveType_(r[0])]);
    sh.getRange(2, typeCol, out.length, 1).setValues(out);
    return { ok: true, version: API_VERSION, updated: out.length, typeCol: typeCol };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* =========================
   Duplicate object builder
========================= */
function apiBuildDuplicateObj_(sh, row, key) {
  const tsInfo    = apiGetTimestampInfoByRow_(sh, row);
  const submitter = apiGetSubmitterNameByRow_(sh, row);
  const dm        = apiGetTextByRowCol_(sh, row, 3);
  const br        = apiGetTextByRowCol_(sh, row, 4);
  const dt        = apiGetTextByRowCol_(sh, row, 5);
  const sl        = apiGetTextByRowCol_(sh, row, 6);
  const type      = apiGetTypeByRow_(sh, row, br);
  return {
    row: row, key: key,
    timestamp: tsInfo.timestamp, timestampText: tsInfo.timestampText,
    submitter_name: submitter, district_manager: dm,
    branch: br, submit_date: dt, submit_time_slot: sl, type: type
  };
}

function apiGetTextByRowCol_(sh, row, col) {
  try {
    const v = sh.getRange(row, col).getValue();
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
      if (col === 5) return Utilities.formatDate(v, API_TIMEZONE, "yyyy-MM-dd");
      return Utilities.formatDate(v, API_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    }
    return String(v || "").trim();
  } catch(_) { return ""; }
}

function apiGetTimestampInfoByRow_(sh, row) {
  try {
    const v = sh.getRange(row, 1).getValue();
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
      return { timestamp: v.getTime(), timestampText: Utilities.formatDate(v, API_TIMEZONE, "yyyy-MM-dd HH:mm:ss") };
    }
    const s = String(v || "").trim();
    return { timestamp: null, timestampText: s };
  } catch (_) { return { timestamp: null, timestampText: "" }; }
}

function apiGetSubmitterNameByRow_(sh, row) {
  try { return String(sh.getRange(row, 2).getValue() || "").trim(); } catch (_) { return ""; }
}

function apiGetTypeByRow_(sh, row, branchFallback) {
  try {
    const typeCol = apiTypeCol_(sh);
    const v       = sh.getRange(row, typeCol).getValue();
    const s       = String(v || "").trim();
    return s || apiResolveType_(branchFallback);
  } catch (_) { return apiResolveType_(branchFallback); }
}

/* =========================
   Canonicalize
========================= */
function apiCanonDate_(s) {
  if (!s) return "";
  if (Object.prototype.toString.call(s) === "[object Date]" && !isNaN(s.getTime())) {
    return Utilities.formatDate(s, API_TIMEZONE, "yyyy-MM-dd");
  }
  if (typeof s === "string") {
    const t = Date.parse(s);
    if (!isNaN(t)) return Utilities.formatDate(new Date(t), API_TIMEZONE, "yyyy-MM-dd");
  }
  s = String(s).trim();
  if (!s) return "";
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    return `${m[3]}-${String(m[2]).padStart(2,"0")}-${String(m[1]).padStart(2,"0")}`;
  }
  return s;
}

function apiCanonSlot_(s) {
  s = String(s || "").trim().replace(":", ".").replace(/\s+/g, "");
  if (!/^\d+(\.\d+)?$/.test(s)) return s;
  const hour = Math.round(Number(s));
  return isFinite(hour) ? String(hour) + ".00" : s;
}

function apiPrevDate_(yyyyMMdd) {
  try {
    const s = String(yyyyMMdd || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
    const d = new Date(s + "T00:00:00+07:00");
    d.setDate(d.getDate() - 1);
    return Utilities.formatDate(d, API_TIMEZONE, "yyyy-MM-dd");
  } catch (_) { return ""; }
}

function apiBranchCode_(br) {
  const s = String(br || "").trim();
  const m = s.match(/^(\d{4,6})/);
  return m ? m[1] : apiNorm_(s);
}

function apiResolveType_(br) {
  const code = apiBranchCode_(br);
  return API_EASY_BRANCH_CODES[code] ? API_TYPE_SANTA_FE_EASY : API_TYPE_SANTA_FE;
}

function apiMakeKey_(dm, br, dt, sl) {
  return [apiNorm_(dm), apiBranchCode_(br), apiCanonDate_(dt), apiCanonSlot_(sl)].join("|");
}

function apiMakeKeyRawSlot_(dm, br, canonDate, rawSlot) {
  return [apiNorm_(dm), apiBranchCode_(br), String(canonDate || "").trim(), String(rawSlot || "").trim()].join("|");
}

/* =========================
   Parse / Utils
========================= */
function apiParseBody_(e) {
  const p   = (e && e.parameter) ? e.parameter : {};
  if (p && Object.keys(p).length) return p;
  const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
  const ct  = (e && e.postData && e.postData.type) ? e.postData.type : "";
  if (ct && ct.indexOf("application/json") >= 0) return raw ? JSON.parse(raw) : {};
  const out = {};
  raw.split("&").forEach(kv => {
    const parts = kv.split("=");
    const k     = parts[0];
    const v     = parts.slice(1).join("=");
    if (!k) return;
    out[decodeURIComponent(k)] = decodeURIComponent((v || "").replace(/\+/g, " "));
  });
  return out;
}

function apiNum_(v)  { const n = Number(String(v ?? "").replace(/,/g, "").trim()); return isFinite(n) ? n : 0; }
function apiInt_(v)  { const n = Number(String(v ?? "").replace(/,/g, "").trim()); return isFinite(n) ? Math.round(n) : 0; }
function apiNorm_(s) { return String(s || "").replace(/\s+/g, " ").trim(); }

/* =========================
   Output helpers
========================= */
function apiJsonOut_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(`${cb}(${out});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function apiHtmlBridgeOut_(obj, reqId) {
  const wire = JSON.stringify({
    source: "santafe-sales-api",
    reqId: String(reqId || ""),
    payload: obj || {}
  }).replace(/</g, "\\u003c");

  const html = [
    '<!doctype html><html><head><meta charset="utf-8"></head><body>',
    '<script>',
    '(function(){',
    'var msg = ' + wire + ';',
    'try { top.postMessage(msg, "*"); } catch (e) {}',
    'try { parent.postMessage(msg, "*"); } catch (e) {}',
    'document.body.innerHTML = "OK";',
    '})();',
    '</script>',
    '</body></html>'
  ].join('');

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
