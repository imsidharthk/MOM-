
const SHEET_AP_UPDATES_LOG = "Action Point Updates Log";


const SHEET_ACTION_MASTER  = "Meeting Action Points Master";     // Tab 3 output
const UPLOAD_FOLDER_ID = "1BY1T3o1lnk7kG4RjL0U6mkk5vASA5PaQ";  // Drive folder for uploads

/** ---------- SHEET NAMES ---------- */
const SHEET_TEAM_MEMBERS   = "Team Members";
const SHEET_MEETING_DETAILS = "Meeting Details";          // Meeting name source
const SHEET_OUTPUT          = "Meeting Details";          // Tab 1 output (you set)
const SHEET_RESULTS_MASTER  = "Meeting Results Master";   // Tab 2 output

/** ---------- doGet ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("MOM 2.0");
}

/** ---------- Helpers ---------- */
function _getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _uniqueClean_(arr) {
  const set = new Set();
  arr.forEach(v => {
    const s = (v ?? "").toString().trim();
    if (s) set.add(s);
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

function _ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    return;
  }
  // If sheet exists but first row empty-ish, force set headers (safe)
  const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const allBlank = existing.every(v => (v ?? "").toString().trim() === "");
  if (allBlank) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  }
}

/** ---------- TAB 1 APIs (same) ---------- */

// Team Members (multi-select)
function getTeamMemberNames() {
  const sh = _getOrCreateSheet_(SHEET_TEAM_MEMBERS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => (h ?? "").toString().trim().toLowerCase());
  const nameIdx = headers.indexOf("name");
  if (nameIdx === -1) return [];

  const values = data.slice(1).map(r => r[nameIdx]);
  return _uniqueClean_(values);
}

// Frequency dropdown
function getFrequencies() {
  const sh = _getOrCreateSheet_(SHEET_TEAM_MEMBERS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => (h ?? "").toString().trim().toLowerCase());
  const freqIdx = headers.indexOf("frequency");
  if (freqIdx === -1) return [];

  const values = data.slice(1).map(r => r[freqIdx]);
  return _uniqueClean_(values);
}

// Submit New Meeting (Tab 1)
function submitNewMeeting(payload) {
  const meetingName = (payload?.meetingName ?? "").toString().trim();
  const department  = (payload?.department ?? "").toString().trim();
  const frequency   = (payload?.frequency ?? "").toString().trim();
  const participantsArr = Array.isArray(payload?.participants) ? payload.participants : [];

  if (!meetingName) return { ok:false, message:"Name of Meeting is required." };
  if (!department)  return { ok:false, message:"Department is required." };
  if (!frequency)   return { ok:false, message:"Frequency is required." };

  const participants = participantsArr
    .map(p => (p ?? "").toString().trim())
    .filter(Boolean)
    .join(", ");

  const sh = _getOrCreateSheet_(SHEET_OUTPUT);

  _ensureHeaders_(sh, ["Meeting Name","Department","List of Participants","Frequency"]);
  sh.appendRow([meetingName, department, participants, frequency]);

  return { ok:true, message:"Meeting saved successfully." };
}

/** ---------- TAB 2: New Result APIs ---------- */

// Meeting names from "Meeting Details" column A, header "Meeting Name"
function getMeetingNames() {
  const sh = _getOrCreateSheet_(SHEET_MEETING_DETAILS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Read column A
  const colA = sh.getRange(1, 1, lastRow, 1).getValues().flat().map(v => (v ?? "").toString().trim());
  const header = (colA[0] || "").toLowerCase();

  // If header exists, use from row 2; else use from row 1
  const startIndex = header === "meeting name" ? 1 : 0;
  const names = colA.slice(startIndex);
  return _uniqueClean_(names);
}

// Named range ResultNo -> next number
function getNextResultNo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName("ResultNo");
  if (!range) return { ok:false, message:'Named range "ResultNo" not found.' };

  const current = Number(range.getValue());
  const safeCurrent = Number.isFinite(current) ? current : 0;
  return { ok:true, next: safeCurrent + 1 };
}

// Submit multiple results (2D array append)
function submitMeetingResults(payload) {
  // payload: { rows: [ {resultNo, meetingName, resultShort, resultLong, targetDate} ] }
  const rows = Array.isArray(payload?.rows) ? payload.rows : [];
  if (rows.length === 0) return { ok:false, message:"No results to submit." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultNoRange = ss.getRangeByName("ResultNo");
  if (!resultNoRange) return { ok:false, message:'Named range "ResultNo" not found.' };

  const sh = _getOrCreateSheet_(SHEET_RESULTS_MASTER);
  _ensureHeaders_(sh, ["Result No","Meeting Name","Result Short","Result Long","Target Date"]);

  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";

  // Build 2D array
  const out = [];
  let maxResultNo = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};
    const resultNo = Number(r.resultNo);
    const meetingName = (r.meetingName ?? "").toString().trim();
    const resultShort = (r.resultShort ?? "").toString().trim();
    const resultLong  = (r.resultLong ?? "").toString().trim();
    const targetDateRaw = (r.targetDate ?? "").toString().trim(); // expect yyyy-mm-dd from input[type=date]

    if (!Number.isFinite(resultNo)) return { ok:false, message:`Invalid Result No at row ${i+1}` };
    if (!meetingName) return { ok:false, message:`Meeting Name required at row ${i+1}` };
    if (!resultShort) return { ok:false, message:`Result Short required at row ${i+1}` };
    if (!resultLong)  return { ok:false, message:`Result Long required at row ${i+1}` };
    if (!targetDateRaw) return { ok:false, message:`Target Date required at row ${i+1}` };

    // 2–4 words validation (server side safety)
    const wordCount = resultShort.split(/\s+/).filter(Boolean).length;
    if (wordCount < 2 || wordCount > 4) {
      return { ok:false, message:`Result Short must be 2–4 words (row ${i+1}).` };
    }

    // Date formatting dd/MM/yyyy
    const d = new Date(targetDateRaw);
    if (isNaN(d.getTime())) return { ok:false, message:`Invalid Target Date at row ${i+1}` };

    const formatted = Utilities.formatDate(d, tz, "dd/MM/yyyy");

    out.push([resultNo, meetingName, resultShort, resultLong, formatted]);
    if (resultNo > maxResultNo) maxResultNo = resultNo;
  }

  // Append 2D array in one shot
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, out.length, out[0].length).setValues(out);

  // Update named range to last used result no
  resultNoRange.setValue(maxResultNo);

  return { ok:true, message:`Submitted ${out.length} result(s) successfully.`, lastResultNo: maxResultNo };
}
/** =========================
 *  TAB 3: DATA SOURCES
 *  ========================= */

// Meeting name unique + Result Short list by meeting from "Meeting Results Master"
function getMeetingResultsMap() {
  const sh = _getOrCreateSheet_(SHEET_RESULTS_MASTER);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { meetingNames: [], resultShortByMeeting: {} };

  const headers = data[0].map(h => (h ?? "").toString().trim().toLowerCase());
  const meetingIdx = headers.indexOf("meeting name");
  const resultShortIdx = headers.indexOf("result short");

  if (meetingIdx === -1 || resultShortIdx === -1) {
    return { meetingNames: [], resultShortByMeeting: {} };
  }

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const meeting = (data[i][meetingIdx] ?? "").toString().trim();
    const rs = (data[i][resultShortIdx] ?? "").toString().trim();
    if (!meeting || !rs) continue;
    if (!map[meeting]) map[meeting] = new Set();
    map[meeting].add(rs);
  }

  const resultShortByMeeting = {};
  Object.keys(map).forEach(m => resultShortByMeeting[m] = Array.from(map[m]).sort((a,b)=>a.localeCompare(b)));

  const meetingNames = Object.keys(resultShortByMeeting).sort((a,b)=>a.localeCompare(b));
  return { meetingNames, resultShortByMeeting };
}

/** =========================
 *  TAB 3: UPLOAD FILE (optional)
 *  ========================= */

// payload: { fileName, mimeType, base64Data } where base64Data is DataURL part after comma
function uploadDocument(payload) {
  const fileName = (payload?.fileName ?? "").toString().trim() || "MOM-Upload";
  const mimeType = (payload?.mimeType ?? "").toString().trim() || "application/octet-stream";
  const base64Data = (payload?.base64Data ?? "").toString();

  if (!base64Data) return { ok: false, message: "No file data." };

  // ✅ Use existing folder by ID
  const folder = DriveApp.getFolderById("1BY1T3o1lnk7kG4RjL0U6mkk5vASA5PaQ");

  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const file = folder.createFile(blob);

  // 🔒 Recommended: keep private (no public link)
  // If you WANT public link, tell me – I’ll change this
  // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    ok: true,
    link: file.getUrl(),   // Drive file link
    id: file.getId()
  };
}


/** =========================
 *  TAB 3: SUBMIT ACTION POINTS
 *  ========================= */

// rows: expanded rows already (1 row per doer)
// Each row: [Meeting name, Result Short, Action point, Task type, Doer, Follower, Frequency, Start date, Target Date, Urgency, Upload link]
function submitActionPointsV2(payload) {
  const rows = Array.isArray(payload?.rows) ? payload.rows : [];
  if (rows.length === 0) return { ok:false, message:"No action points to submit." };

  const sh = _getOrCreateSheet_(SHEET_ACTION_MASTER);

  _ensureHeaders_(sh, [
    "ActionNo",
    "Meeting name",
    "Result Short",
    "Action point",
    "Task type",
    "Doer",
    "Follower",
    "Frequency",
    "Start date",
    "Target Date",
    "Urgency",
    "Upload any document link"
  ]);

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  return { ok:true, message:`Submitted ${rows.length} row(s) successfully.` };
}

function getNextActionNoV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName("VActionNo");
  if (!range) return { ok:false, message:'Named range "VActionNo" not found.' };

  const current = Number(range.getValue());
  const safeCurrent = Number.isFinite(current) ? current : 0;
  return { ok:true, next: safeCurrent + 1 };
}
function _headerIndexMap_(headersRow) {
  const m = {};
  headersRow.forEach((h, i) => {
    const key = (h ?? "").toString().trim().toLowerCase();
    if (key) m[key] = i;
  });
  return m;
}

function _parseDateFlexible_(v) {
  if (!v) return null;

  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = v.toString().trim();
  if (!s) return null;

  // yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // dd/MM/yyyy
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]) - 1, yy = Number(m[3]);
    const d = new Date(yy, mm, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function getPendingActionUpdateData() {
  const sh = _getOrCreateSheet_(SHEET_ACTION_MASTER);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { meetingNames: [], resultShortByMeeting: {}, actionsByMeetingResult: {} };

  const headers = data[0].map(h => (h ?? "").toString().trim().toLowerCase());

  const idxActionNo   = headers.indexOf("actionno");
  const idxMeeting    = headers.indexOf("meeting name");
  const idxResult     = headers.indexOf("result short");
  const idxAction     = headers.indexOf("action point");
  const idxTaskType   = headers.indexOf("task type");
  const idxDoer       = headers.indexOf("doer");
  const idxStart      = headers.indexOf("start date");
  const idxTarget     = headers.indexOf("target date");
  const idxStatus     = headers.indexOf("status");

  // Basic safety
  if ([idxActionNo, idxMeeting, idxResult, idxAction, idxTaskType, idxDoer, idxStart, idxTarget, idxStatus].some(x => x === -1)) {
    return { meetingNames: [], resultShortByMeeting: {}, actionsByMeetingResult: {} };
  }

  // Pending definition: status blank OR "Pending" (case-insensitive)
  function isPending(st) {
    const s = (st ?? "").toString().trim().toLowerCase();
    return s === "" || s === "pending";
  }

  const meetingSet = new Set();
  const resultMap = {}; // meeting -> Set(resultShort)
  const actionsMap = {}; // "meeting||result" -> array of {key,label, taskType, targetDateISO}

  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    const actionNo = row[idxActionNo];
    const meeting  = (row[idxMeeting] ?? "").toString().trim();
    const resultS  = (row[idxResult] ?? "").toString().trim();
    const actionP  = (row[idxAction] ?? "").toString().trim();
    const taskType = (row[idxTaskType] ?? "").toString().trim();
    const doer     = (row[idxDoer] ?? "").toString().trim();
    const status   = row[idxStatus];
    const targetDt = row[idxTarget];

    if (!meeting || !resultS || !actionP) continue;
    if (!isPending(status)) continue; // only pending/blank

    meetingSet.add(meeting);

    if (!resultMap[meeting]) resultMap[meeting] = new Set();
    resultMap[meeting].add(resultS);

    const mrKey = meeting + "||" + resultS;
    if (!actionsMap[mrKey]) actionsMap[mrKey] = [];

    // Unique row key: ActionNo + Doer (so one exact row updates)
    const key = String(actionNo) + "||" + doer;

    // Convert target date to ISO (yyyy-mm-dd) if it’s a Date
    let targetISO = "";
    if (targetDt instanceof Date && !isNaN(targetDt.getTime())) {
      targetISO = Utilities.formatDate(targetDt, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
    } else if (typeof targetDt === "string") {
      // if stored as yyyy-mm-dd already, keep it if looks like ISO
      targetISO = targetDt.trim();
    }

    const label = `#${actionNo} | ${doer} | ${actionP}`;

    actionsMap[mrKey].push({
      key,
      label,
      actionNo: String(actionNo),
      doer,
      actionPoint: actionP,
      taskType,
      targetDateISO: targetISO
    });
  }

  const meetingNames = Array.from(meetingSet).sort((a,b)=>a.localeCompare(b));

  const resultShortByMeeting = {};
  Object.keys(resultMap).forEach(m => {
    resultShortByMeeting[m] = Array.from(resultMap[m]).sort((a,b)=>a.localeCompare(b));
  });

  // Sort actions list by actionNo (numeric)
  Object.keys(actionsMap).forEach(k => {
    actionsMap[k].sort((x,y) => Number(x.actionNo) - Number(y.actionNo));
  });

  return { meetingNames, resultShortByMeeting, actionsByMeetingResult: actionsMap };
}
function updateActionPoint(payload) {
  // payload: { key:"ActionNo||Doer", status, remarks, nextTargetDate, nextFollowUpDate }
  const key = (payload?.key ?? "").toString();
  const status = (payload?.status ?? "").toString().trim();
  const remarks = (payload?.remarks ?? "").toString().trim();
  const nextTargetDate = (payload?.nextTargetDate ?? "").toString().trim();     // yyyy-mm-dd or ""
  const nextFollowUpDate = (payload?.nextFollowUpDate ?? "").toString().trim(); // yyyy-mm-dd or ""

  if (!key) return { ok:false, message:"No action selected." };
  if (!status) return { ok:false, message:"Status is required." };

  const [actionNoStr, doer] = key.split("||");
  if (!actionNoStr || !doer) return { ok:false, message:"Invalid action key." };

  const sh = _getOrCreateSheet_(SHEET_ACTION_MASTER);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok:false, message:"No data found in sheet." };

  const headers = data[0].map(h => (h ?? "").toString().trim().toLowerCase());
  const idxActionNo = headers.indexOf("actionno");
  const idxDoer     = headers.indexOf("doer");
  const idxStart    = headers.indexOf("start date");
  const idxTarget   = headers.indexOf("target date");
  const idxStatus   = headers.indexOf("status");
  const idxRemarks  = headers.indexOf("remarks");

  if ([idxActionNo, idxDoer, idxStart, idxTarget, idxStatus, idxRemarks].some(x => x === -1)) {
    return { ok:false, message:"Required columns not found (ActionNo/Doer/Start date/Target Date/Status/Remarks)." };
  }

  // Find row
  let foundRow = -1;
  for (let r = 1; r < data.length; r++) {
    const aNo = String(data[r][idxActionNo]);
    const d = (data[r][idxDoer] ?? "").toString().trim();
    if (aNo === String(actionNoStr) && d === doer) {
      foundRow = r + 1; // sheet row number
      break;
    }
  }

  if (foundRow === -1) return { ok:false, message:"Action row not found for selected ActionNo + Doer." };

  // Update only selected fields
  sh.getRange(foundRow, idxStatus + 1).setValue(status);

  if (remarks) {
    sh.getRange(foundRow, idxRemarks + 1).setValue(remarks);
  }

  if (nextTargetDate) {
    sh.getRange(foundRow, idxTarget + 1).setValue(nextTargetDate);
  }

  if (nextFollowUpDate) {
    sh.getRange(foundRow, idxStart + 1).setValue(nextFollowUpDate);
  }

  SpreadsheetApp.flush();
  return { ok:true, message:`Updated ActionNo ${actionNoStr} (${doer}) successfully.` };
}

