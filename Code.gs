// ════════════════════════════════════════════════════════════
//  Santa Fe Sales Report — Google Apps Script Backend
//  Deploy: Execute as "Me", Access "Anyone"
//  Sheet: "Sales"   → ข้อมูลยอดขาย
//         "Plan"    → ข้อมูล Plan Sale รายวัน
// ════════════════════════════════════════════════════════════

var SALES_HDR = [
  "timestamp","submitter_name","district_manager","branch","branch_code",
  "submit_date","submit_time_slot",
  "plan_sale","actual_sale","sale_dine_in","sale_take_away","sale_grab",
  "sale_lineman","sale_shopeefood",
  "total_trans","trans_dine_in","trans_take_away","trans_grab",
  "trans_lineman","trans_shopeefood","customer","labour_hour","labour_baht",
  "edit_count","last_edited"
];

var PLAN_HDR = ["branch_code","date","plan_sale","submitter_name","updated_at"];

// ── Entry points ─────────────────────────────────────────────
// doGet: GET request ไม่มี redirect → CORS ทำงานได้ทันที (ใช้หลัก)
function doGet(e) {
  return handleRequest(e);
}

// doPost: เก็บไว้เป็น fallback
function doPost(e) {
  return handleRequest(e);
}

var WRITE_MODES = { submit: true, edit: true, savePlan: true };

function handleRequest(e) {
  try {
    var p    = e.parameter;
    var mode = p.mode || "";
    var payload;

    if (WRITE_MODES[mode]) {
      // ล็อก script ป้องกัน race condition ตอนหลายสาขาส่ง/แก้ไขพร้อมกัน
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(12000); // รอสูงสุด 12 วินาที
      } catch (e) {
        return respond({ ok: false, error: "LOCK_TIMEOUT", userMessage: "ระบบยุ่งอยู่ กรุณาลองใหม่อีกครั้ง" });
      }
      try {
        if      (mode === "submit")   payload = handleSubmit(p);
        else if (mode === "savePlan") payload = handleSavePlan(p);
        else if (mode === "edit")     payload = handleEdit(p);
      } finally {
        lock.releaseLock();
      }
    } else {
      if      (mode === "getPlan")  payload = handleGetPlan(p);
      else if (mode === "history")  payload = handleHistory(p);
      else                          payload = { ok: false, error: "Unknown mode: " + mode };
    }

    return respond(payload);
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ────────────────────────────────────────────
function getSalesSheet() { return getOrCreate("Sales", SALES_HDR); }
function getPlanSheet()  { return getOrCreate("Plan",  PLAN_HDR);  }

function getOrCreate(name, headers) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

// ── Type helpers ─────────────────────────────────────────────
// Google Sheets auto-converts "5001" → 5001 (number) and
// "2026-04-23" → Date object. Always use these helpers when
// reading cell values back for comparison or JSON output.

function toDateStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd");
  }
  return String(v);
}

// Sheets แปลง "16.00" → number 16 → ต้อง toFixed(2) คืนกลับ
function toSlotStr(v) {
  if (typeof v === "number") return v.toFixed(2);
  return String(v);
}

function toNum(v) {
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

// ── Plan sync helper ─────────────────────────────────────────
// เรียกหลัง submit/edit ที่มี plan_sale > 0 เพื่อ sync ไป Plan sheet
function upsertPlan(branch_code, date, plan_sale, submitter_name) {
  var sheet = getPlanSheet();
  var data  = sheet.getDataRange().getValues();
  var now   = new Date().toISOString();

  for (var i = 1; i < data.length; i++) {
    var rowCode = String(data[i][0]).replace(/^'/, "");
    var rowDate = toDateStr(data[i][1]);
    if (rowCode === String(branch_code) && rowDate === String(date)) {
      sheet.getRange(i + 1, 3).setValue(plan_sale);
      sheet.getRange(i + 1, 4).setValue(submitter_name || "");
      sheet.getRange(i + 1, 5).setValue(now);
      return;
    }
  }
  sheet.appendRow(["'" + branch_code, date, plan_sale, submitter_name || "", now]);
}

// ════════════════════════════════════════════════════════════
//  MODE: submit
// ════════════════════════════════════════════════════════════
function handleSubmit(p) {
  var sheet = getSalesSheet();
  var data  = sheet.getDataRange().getValues();

  var has1600 = false;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rCode = String(r[4]).replace(/^'/, "");
    var rDate = toDateStr(r[5]);
    var rSlot = toSlotStr(r[6]);

    // ตรวจ duplicate
    if (rCode === String(p.branch_code) &&
        rDate === String(p.submit_date) &&
        rSlot === String(p.submit_time_slot)) {
      return {
        ok: false,
        duplicate: {
          submit_date:      rDate,
          submit_time_slot: rSlot,
          submitter_name:   String(r[1])
        }
      };
    }

    // ตรวจว่ามี 16.00 แล้ว (สำหรับ check สิ้นวัน)
    if (rCode === String(p.branch_code) &&
        rDate === String(p.submit_date) &&
        rSlot === "16.00") {
      has1600 = true;
    }
  }

  // สิ้นวัน ต้องส่ง 16.00 ก่อน
  if (p.submit_time_slot === "สิ้นวัน" && !has1600) {
    return { ok: false, userMessage: "กรุณาส่งข้อมูลรอบ 16.00 ก่อนส่งรอบสิ้นวัน" };
  }

  sheet.appendRow([
    new Date().toISOString(),
    p.submitter_name   || "",
    p.district_manager || "",
    p.branch           || "",
    "'" + (p.branch_code || ""),   // prefix ' ป้องกัน Sheets แปลงเป็น number
    p.submit_date      || "",
    p.submit_time_slot || "",
    toNum(p.plan_sale),      toNum(p.actual_sale),
    toNum(p.sale_dine_in),   toNum(p.sale_take_away),
    toNum(p.sale_grab),      toNum(p.sale_lineman),
    toNum(p.sale_shopeefood),
    toNum(p.total_trans),    toNum(p.trans_dine_in),
    toNum(p.trans_take_away),toNum(p.trans_grab),
    toNum(p.trans_lineman),  toNum(p.trans_shopeefood),
    toNum(p.customer),       toNum(p.labour_hour),
    toNum(p.labour_baht),
    0, ""
  ]);

  // sync Plan Sale → Plan sheet (เฉพาะเมื่อมีค่า plan_sale)
  if (toNum(p.plan_sale) > 0) {
    upsertPlan(p.branch_code, p.submit_date, toNum(p.plan_sale), p.submitter_name);
  }

  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  MODE: getPlan
// ════════════════════════════════════════════════════════════
function handleGetPlan(p) {
  var sheet = getPlanSheet();
  var data  = sheet.getDataRange().getValues();
  var rows  = [];

  for (var i = 1; i < data.length; i++) {
    var r       = data[i];
    var rowCode = String(r[0]).replace(/^'/, "");  // ลบ ' prefix ถ้ามี
    var rowDate = toDateStr(r[1]);

    if (rowCode !== String(p.branch_code)) continue;

    if (p.date) {
      if (rowDate === String(p.date)) rows.push({ date: rowDate, plan_sale: r[2] });
    } else if (p.year_month) {
      if (rowDate.indexOf(String(p.year_month)) === 0) rows.push({ date: rowDate, plan_sale: r[2] });
    }
  }

  return { ok: true, rows: rows };
}

// ════════════════════════════════════════════════════════════
//  MODE: savePlan
// ════════════════════════════════════════════════════════════
function handleSavePlan(p) {
  var sheet = getPlanSheet();
  var data  = sheet.getDataRange().getValues();

  var entries;
  try { entries = JSON.parse(p.entries || "[]"); }
  catch (e) { return { ok: false, error: "Invalid entries JSON" }; }

  var saved = 0;
  var now   = new Date().toISOString();

  entries.forEach(function(entry) {
    if (!entry.date) return;
    var found = false;

    for (var i = 1; i < data.length; i++) {
      var rowCode = String(data[i][0]).replace(/^'/, "");
      var rowDate = toDateStr(data[i][1]);
      if (rowCode === String(p.branch_code) && rowDate === String(entry.date)) {
        sheet.getRange(i + 1, 3).setValue(entry.plan_sale || 0);
        sheet.getRange(i + 1, 4).setValue(p.submitter_name || "");
        sheet.getRange(i + 1, 5).setValue(now);
        data[i][2] = entry.plan_sale || 0;
        found = true;
        saved++;
        break;
      }
    }

    if (!found) {
      var newRow = ["'" + p.branch_code, entry.date, entry.plan_sale || 0, p.submitter_name || "", now];
      sheet.appendRow(newRow);
      data.push(newRow);
      saved++;
    }
  });

  return { ok: true, saved: saved };
}

// ════════════════════════════════════════════════════════════
//  MODE: history
// ════════════════════════════════════════════════════════════
function handleHistory(p) {
  var sheet     = getSalesSheet();
  var data      = sheet.getDataRange().getValues();
  var days      = parseInt(p.days) || 30;
  var cutoff    = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  var cutoffStr = Utilities.formatDate(cutoff, "Asia/Bangkok", "yyyy-MM-dd");

  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r       = data[i];
    var rowCode = String(r[4]).replace(/^'/, "");
    var rowDate = toDateStr(r[5]);

    if (rowCode !== String(p.branch_code)) continue;
    if (rowDate < cutoffStr) continue;
    rows.push(rowToObj(r, i + 1));
  }

  rows.sort(function(a, b) {
    if (a.submit_date !== b.submit_date) return b.submit_date.localeCompare(a.submit_date);
    if (a.submit_time_slot === "สิ้นวัน") return 1;
    if (b.submit_time_slot === "สิ้นวัน") return -1;
    return 0;
  });

  return { ok: true, rows: rows };
}

// ════════════════════════════════════════════════════════════
//  MODE: edit
// ════════════════════════════════════════════════════════════
function handleEdit(p) {
  var sheet  = getSalesSheet();
  var rowNum = parseInt(p._row);

  if (!rowNum || rowNum < 2) return { ok: false, error: "Invalid row number" };

  var existing = sheet.getRange(rowNum, 1, 1, SALES_HDR.length).getValues()[0];
  var rowCode  = String(existing[4]).replace(/^'/, "");

  if (rowCode !== String(p.branch_code)) {
    return { ok: false, userMessage: "ไม่สามารถแก้ไขข้อมูลสาขาอื่นได้" };
  }

  var editCount = (parseInt(existing[23]) || 0) + 1;

  sheet.getRange(rowNum, 1, 1, SALES_HDR.length).setValues([[
    existing[0],
    p.submitter_name   || existing[1],
    p.district_manager || existing[2],
    p.branch           || existing[3],
    existing[4],                         // branch_code เดิม (ไม่แตะ format)
    existing[5],                         // submit_date เดิม
    p.submit_time_slot || existing[6],
    toNum(p.plan_sale),      toNum(p.actual_sale),
    toNum(p.sale_dine_in),   toNum(p.sale_take_away),
    toNum(p.sale_grab),      toNum(p.sale_lineman),
    toNum(p.sale_shopeefood),
    toNum(p.total_trans),    toNum(p.trans_dine_in),
    toNum(p.trans_take_away),toNum(p.trans_grab),
    toNum(p.trans_lineman),  toNum(p.trans_shopeefood),
    toNum(p.customer),       toNum(p.labour_hour),
    toNum(p.labour_baht),
    editCount,
    new Date().toISOString()
  ]]);

  var newPlan    = toNum(p.plan_sale);
  var targetDate = toDateStr(existing[5]);

  if (newPlan > 0) {
    // sync Plan Sale → Plan sheet
    upsertPlan(p.branch_code, targetDate, newPlan, p.submitter_name);

    // sync plan_sale ไปยัง row อื่นของวันเดียวกัน (รอบ 16.00 ↔ สิ้นวัน)
    var allRows = sheet.getDataRange().getValues();
    for (var j = 1; j < allRows.length; j++) {
      if (j + 1 === rowNum) continue;
      var rd     = allRows[j];
      var rdCode = String(rd[4]).replace(/^'/, "");
      var rdDate = toDateStr(rd[5]);
      if (rdCode === String(p.branch_code) && rdDate === targetDate) {
        sheet.getRange(j + 1, 8).setValue(newPlan); // col 8 = plan_sale
      }
    }
  }

  return { ok: true, edited: true, editCount: editCount };
}

// ── Row → object (แปลง Date/number กลับเป็น string ก่อนส่ง) ─
function rowToObj(r, rowNum) {
  return {
    _row:             rowNum,
    timestamp:        String(r[0]),
    submitter_name:   String(r[1]),
    district_manager: String(r[2]),
    branch:           String(r[3]),
    branch_code:      String(r[4]).replace(/^'/, ""),
    submit_date:      toDateStr(r[5]),
    submit_time_slot: toSlotStr(r[6]),
    plan_sale:        r[7],
    actual_sale:      r[8],
    sale_dine_in:     r[9],
    sale_take_away:   r[10],
    sale_grab:        r[11],
    sale_lineman:     r[12],
    sale_shopeefood:  r[13],
    total_trans:      r[14],
    trans_dine_in:    r[15],
    trans_take_away:  r[16],
    trans_grab:       r[17],
    trans_lineman:    r[18],
    trans_shopeefood: r[19],
    customer:         r[20],
    labour_hour:      r[21],
    labour_baht:      r[22],
    edit_count:       r[23],
    last_edited:      String(r[24] || "")
  };
}
