// ════════════════════════════════════════════════════════════
//  Santa Fe Sales Report — Google Apps Script Backend
//  Deploy: Execute as "Me", Access "Anyone"
//  Sheet: "Sales"   → ข้อมูลยอดขาย
//         "Plan"    → ข้อมูล Plan Sale รายวัน
// ════════════════════════════════════════════════════════════

// ── Column headers ──────────────────────────────────────────
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

// ── Entry point ──────────────────────────────────────────────
function doPost(e) {
  try {
    var p    = e.parameter;
    var mode = p.mode || "";
    var payload;

    if      (mode === "submit")   payload = handleSubmit(p);
    else if (mode === "getPlan")  payload = handleGetPlan(p);
    else if (mode === "savePlan") payload = handleSavePlan(p);
    else if (mode === "history")  payload = handleHistory(p);
    else if (mode === "edit")     payload = handleEdit(p);
    else                          payload = { ok: false, error: "Unknown mode: " + mode };

    return respond(payload);

  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

// ── Return JSON directly (no iframe/postMessage needed) ──────
function respond(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Lazy-init sheets ─────────────────────────────────────────
function getSalesSheet() { return getOrCreate("Sales",  SALES_HDR); }
function getPlanSheet()  { return getOrCreate("Plan",   PLAN_HDR);  }

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

// ════════════════════════════════════════════════════════════
//  MODE: submit — บันทึกยอดขายใหม่
// ════════════════════════════════════════════════════════════
function handleSubmit(p) {
  var sheet = getSalesSheet();
  var data  = sheet.getDataRange().getValues();

  // ตรวจซ้ำ (branch_code + submit_date + submit_time_slot)
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[4] === p.branch_code &&
        r[5] === p.submit_date &&
        r[6] === p.submit_time_slot) {
      return {
        ok: false,
        duplicate: {
          submit_date:      r[5],
          submit_time_slot: r[6],
          submitter_name:   r[1]
        }
      };
    }
  }

  sheet.appendRow([
    new Date().toISOString(),
    p.submitter_name   || "",
    p.district_manager || "",
    p.branch           || "",
    p.branch_code      || "",
    p.submit_date      || "",
    p.submit_time_slot || "",
    toNum(p.plan_sale),     toNum(p.actual_sale),
    toNum(p.sale_dine_in),  toNum(p.sale_take_away),
    toNum(p.sale_grab),     toNum(p.sale_lineman),
    toNum(p.sale_shopeefood),
    toNum(p.total_trans),   toNum(p.trans_dine_in),
    toNum(p.trans_take_away),toNum(p.trans_grab),
    toNum(p.trans_lineman), toNum(p.trans_shopeefood),
    toNum(p.customer),      toNum(p.labour_hour),
    toNum(p.labour_baht),
    0,   // edit_count
    ""   // last_edited
  ]);

  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  MODE: getPlan — ดึง Plan Sale รายวัน
//  params: branch_code + (year_month หรือ date)
// ════════════════════════════════════════════════════════════
function handleGetPlan(p) {
  var sheet = getPlanSheet();
  var data  = sheet.getDataRange().getValues();
  var rows  = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[0] !== p.branch_code) continue;
    var dateStr = String(r[1]);

    if (p.date) {
      if (dateStr === p.date) rows.push({ date: dateStr, plan_sale: r[2] });
    } else if (p.year_month) {
      if (dateStr.indexOf(p.year_month) === 0) rows.push({ date: dateStr, plan_sale: r[2] });
    }
  }

  return { ok: true, rows: rows };
}

// ════════════════════════════════════════════════════════════
//  MODE: savePlan — บันทึก/อัปเดต Plan Sale ทั้งเดือน
//  params: branch_code, submitter_name, entries (JSON string)
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
      if (data[i][0] === p.branch_code && String(data[i][1]) === String(entry.date)) {
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
      var newRow = [p.branch_code, entry.date, entry.plan_sale || 0, p.submitter_name || "", now];
      sheet.appendRow(newRow);
      data.push(newRow);
      saved++;
    }
  });

  return { ok: true, saved: saved };
}

// ════════════════════════════════════════════════════════════
//  MODE: history — ดูประวัติ N วันล่าสุด
//  params: branch_code, days
// ════════════════════════════════════════════════════════════
function handleHistory(p) {
  var sheet   = getSalesSheet();
  var data    = sheet.getDataRange().getValues();
  var days    = parseInt(p.days) || 30;

  var cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  var cutoffStr = cutoff.toISOString().slice(0, 10);

  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[4] !== p.branch_code) continue;
    if (String(r[5]) < cutoffStr)  continue;
    rows.push(rowToObj(r, i + 1));
  }

  // เรียงวันที่ใหม่ก่อน, สิ้นวันอยู่หลัง 16.00 ในวันเดียวกัน
  rows.sort(function(a, b) {
    if (a.submit_date !== b.submit_date) return b.submit_date.localeCompare(a.submit_date);
    if (a.submit_time_slot === "สิ้นวัน") return 1;
    if (b.submit_time_slot === "สิ้นวัน") return -1;
    return 0;
  });

  return { ok: true, rows: rows };
}

// ════════════════════════════════════════════════════════════
//  MODE: edit — แก้ไขแถวที่ระบุ
//  params: _row, branch_code, + ทุก numeric fields
// ════════════════════════════════════════════════════════════
function handleEdit(p) {
  var sheet  = getSalesSheet();
  var rowNum = parseInt(p._row);

  if (!rowNum || rowNum < 2) return { ok: false, error: "Invalid row number" };

  var existing = sheet.getRange(rowNum, 1, 1, SALES_HDR.length).getValues()[0];

  // ตรวจสิทธิ์: branch_code ต้องตรงกัน
  if (String(existing[4]) !== String(p.branch_code)) {
    return { ok: false, userMessage: "ไม่สามารถแก้ไขข้อมูลสาขาอื่นได้" };
  }

  var editCount = (parseInt(existing[23]) || 0) + 1;

  sheet.getRange(rowNum, 1, 1, SALES_HDR.length).setValues([[
    existing[0],              // timestamp เดิม
    p.submitter_name   || existing[1],
    p.district_manager || existing[2],
    p.branch           || existing[3],
    p.branch_code      || existing[4],
    p.submit_date      || existing[5],
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

  return { ok: true, edited: true, editCount: editCount };
}

// ── Helpers ──────────────────────────────────────────────────
function rowToObj(r, rowNum) {
  return {
    _row:             rowNum,
    timestamp:        r[0],
    submitter_name:   r[1],
    district_manager: r[2],
    branch:           r[3],
    branch_code:      r[4],
    submit_date:      r[5],
    submit_time_slot: r[6],
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
    last_edited:      r[24]
  };
}

function toNum(v) {
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
