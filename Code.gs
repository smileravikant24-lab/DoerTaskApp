const MASTER_SHEET_NAME = "Master";
const DOER_LIST_SHEET_NAME = "Doer List";
const ADMIN_EMAIL = "pc@satijapaper.com";
const CLIENT_ID = "609136108830-jf0fovqkjd3l5bv4kal7kg7su2emkfi0.apps.googleusercontent.com";

const STATUS_SUBMITTED = "Submitted";
const STATUS_APPROVED = "Approved";
const STATUS_REJECTED = "Rejected";

// Header names this script will accept for the reject-remark column, in
// order of preference. Whichever one actually exists in the Master sheet's
// header row is used — this used to be hardcoded to "remark" only, so a
// sheet with e.g. "Reject Reason" as the header silently dropped every
// decline reason instead of saving it.
const REMARK_COLUMN_CANDIDATES = ["remark", "reject reason", "rejection reason", "decline reason", "reason for rejection"];

const OVERDUE_LOOKBACK_DAYS = 14;


const PERFORMANCE_LOOKBACK_DAYS = 30;


function doGet(e) {
  try {
    const action = e.parameter.action;
    const idToken = e.parameter.idToken;
    const email = verifyIdToken_(idToken); // throws agar invalid/unauthorized

    let data;
    switch (action) {
      case "dashboard":
        data = getMyDashboard_(email);
        break;
      case "markDone":
        data = markTaskDone_(email, Number(e.parameter.row));
        break;
      case "pendingApprovals":
        assertAdmin_(email);
        data = getPendingApprovals_();
        break;
      case "overdueSummary":
        assertAdmin_(email);
        data = getOverdueSummary_();
        break;
      case "rejectedSummary":
        assertAdmin_(email);
        data = getRejectedSummary_();
        break;
      case "performance":
        assertAdmin_(email);
        data = getUserPerformance_();
        break;
      case "approve":
        assertAdmin_(email);
        data = approveTask_(Number(e.parameter.row));
        break;
      case "reject":
        assertAdmin_(email);
        data = rejectTask_(Number(e.parameter.row), e.parameter.remark || "");
        break;
      case "setDelayReason":
        assertAdmin_(email);
        data = setDelayReason_(Number(e.parameter.row), e.parameter.reason || "");
        break;
      default:
        throw new Error("Unknown action: " + action);
    }
    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// AUTH — Google ID token ko verify karta hai (bina kisi library
// ke, Google ke public tokeninfo endpoint se)
// ------------------------------------------------------------
function verifyIdToken_(idToken) {
  if (!idToken) throw new Error("Missing idToken");

  // Speed optimization: cache the verified email per idToken, so repeated
  // actions (mark done, approve, reject) in the same sign-in session don't
  // each pay for a fresh network round-trip to Google.
  const cache = CacheService.getScriptCache();
  const cacheKey = "idtok_" + hashString_(idToken);
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const resp = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  const info = JSON.parse(resp.getContentText());

  if (info.error) throw new Error("Invalid token: " + (info.error_description || info.error));
  if (info.aud !== CLIENT_ID) throw new Error("Token audience mismatch — galat Client ID.");
  if (!info.email || (info.email_verified !== "true" && info.email_verified !== true)) {
    throw new Error("Email is not verified.");
  }

  const email = info.email;
  const isAdmin = email.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  if (!isAdmin && !isKnownDoer_(email)) {
    throw new Error("This email (" + email + ") was not found in the Doer List.");
  }

  // Cache for the token's remaining validity (capped at 25 min per call,
  // CacheService max) so the very next actions are near-instant.
  const ttlSeconds = info.exp
    ? Math.max(60, Math.min(1500, Number(info.exp) - Math.floor(Date.now() / 1000) - 30))
    : 300;
  cache.put(cacheKey, email, ttlSeconds);

  return email;
}

function hashString_(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(b => ((b < 0 ? b + 256 : b).toString(16).padStart(2, "0"))).join("");
}

function assertAdmin_(email) {
  if (email.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
    throw new Error("Only the admin (" + ADMIN_EMAIL + ") can perform this action.");
  }
}

// ------------------------------------------------------------
// SHEET HELPERS
// ------------------------------------------------------------
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
}

function getHeaderMap_(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headerRow.forEach((h, i) => { if (h) map[String(h).trim().toLowerCase()] = i; });
  return map;
}

function col_(map, name, fallback) {
  const key = name.toLowerCase();
  return (key in map) ? map[key] : fallback;
}

// Returns the 0-based column index for the first header in `candidates`
// that actually exists in `map`, or undefined if none of them do.
function firstMatchingCol_(map, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const key = candidates[i].toLowerCase();
    if (key in map) return map[key];
  }
  return undefined;
}

function formatDateDDMMYYYY_(date) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, "dd/MM/yyyy");
}

function cellToDateString_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") return formatDateDDMMYYYY_(value);
  return String(value).trim();
}

function isKnownDoer_(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "doerlist";
  let emails;
  const cached = cache.get(cacheKey);
  if (cached) {
    emails = JSON.parse(cached);
  } else {
    const sheet = getSheet_(DOER_LIST_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const map = getHeaderMap_(sheet);
    const emailCol = col_(map, "email address", 2);
    emails = [];
    for (let r = 1; r < data.length; r++) {
      const e = String(data[r][emailCol] || "").trim().toLowerCase();
      if (e) emails.push(e);
    }
    cache.put(cacheKey, JSON.stringify(emails), 300); // 5 min — Doer List rarely changes
  }
  return emails.indexOf(email.toLowerCase()) !== -1;
}

// ------------------------------------------------------------
// DOER-SIDE
// ------------------------------------------------------------
function getMyDashboard_(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "dash_" + email.toLowerCase();
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const sheet = getSheet_(MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const cName = col_(map, "option", 0), cEmail = col_(map, "email", 1),
        cDept = col_(map, "department", 2), cFreq = col_(map, "freq", 4),
        cTask = col_(map, "task", 5), cPlanned = col_(map, "planned", 6),
        cActual = col_(map, "actual", 7), cStatus = col_(map, "app status", 11);
  const cReason = map["delay reason"];

  const today = formatDateDDMMYYYY_(new Date());
  const todayMidnight = new Date(); todayMidnight.setHours(0, 0, 0, 0);
  const lookbackStart = new Date(todayMidnight); lookbackStart.setDate(lookbackStart.getDate() - OVERDUE_LOOKBACK_DAYS);
  let name = "";
  const todayTasks = [], rejectedTasks = [], overdueTasks = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[cEmail] || "").trim().toLowerCase() !== email.toLowerCase()) continue;
    if (!name) name = row[cName];

    const status = String(row[cStatus] || "").trim();
    const plannedRaw = row[cPlanned];
    const plannedStr = cellToDateString_(plannedRaw);
    const plannedDate = (Object.prototype.toString.call(plannedRaw) === "[object Date]") ? plannedRaw : null;
    const item = {
      row: r + 1, task: row[cTask], department: row[cDept], freq: row[cFreq],
      planned: plannedStr, actual: cellToDateString_(row[cActual]), status: status
    };

    if (status === STATUS_REJECTED) { rejectedTasks.push(item); continue; }
    if (item.actual) continue; // already done, nothing more to show
    if (plannedStr === today) todayTasks.push(item);
    else if (item.freq !== "D" && plannedDate && plannedDate < todayMidnight && plannedDate >= lookbackStart) {
      item._ts = plannedDate.getTime();
      item.delayReason = (cReason !== undefined ? (row[cReason] || "") : "");
      overdueTasks.push(item);
    }
  }

  overdueTasks.sort((a, b) => a._ts - b._ts); // oldest missed first, for day-wise grouping
  overdueTasks.forEach(t => delete t._ts);
  const result = { name: name, email: email, today: todayTasks, rejected: rejectedTasks, overdue: overdueTasks };
  cache.put(cacheKey, JSON.stringify(result), 30); // 30s — short enough to stay fresh, long enough to skip repeat scans
  return result;
}

function markTaskDone_(email, rowNumber) {
  const sheet = getSheet_(MASTER_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const cEmail = col_(map, "email", 1), cActual = col_(map, "actual", 7),
        cStatus = col_(map, "app status", 11), cFreq = col_(map, "freq", 4);

  const rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowEmail = String(rowValues[cEmail] || "").trim().toLowerCase();
  if (rowEmail !== email.toLowerCase()) throw new Error("This task does not belong to you.");

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");

  // PC (admin) apne khud ke tasks ko approve karne ki zaroorat nahi, aur
  // Daily-frequency tasks bhi kisi approval ke bina seedha Approved ho
  // jaate hain — sirf non-daily doer tasks hi pending-approval me jaate
  // hain. Freq sheet se hi padha jaata hai (client input trust nahi karte).
  const freq = String(rowValues[cFreq] || "").trim();
  const isSelfApprove = email.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  const isDaily = freq === "D";
  const autoApproved = isSelfApprove || isDaily;

  sheet.getRange(rowNumber, cActual + 1).setValue(now);
  sheet.getRange(rowNumber, cStatus + 1).setValue(autoApproved ? STATUS_APPROVED : STATUS_SUBMITTED);

  const cache = CacheService.getScriptCache();
  cache.remove("dash_" + email.toLowerCase());
  cache.remove("pendingApprovals");
  cache.remove("overdueSummary");

  return { row: rowNumber, autoApproved: autoApproved }; // no full re-scan — frontend updates its own list
}

// ------------------------------------------------------------
// ADMIN-SIDE
// ------------------------------------------------------------
function getPendingApprovals_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("pendingApprovals");
  if (cached) return JSON.parse(cached);

  const sheet = getSheet_(MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const cName = col_(map, "option", 0), cEmail = col_(map, "email", 1), cDept = col_(map, "department", 2),
        cFreq = col_(map, "freq", 4), cTask = col_(map, "task", 5), cPlanned = col_(map, "planned", 6),
        cActual = col_(map, "actual", 7), cStatus = col_(map, "app status", 11);

  const pending = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[cStatus] || "").trim() === STATUS_SUBMITTED) {
      pending.push({
        row: r + 1, name: row[cName], email: row[cEmail], department: row[cDept],
        freq: row[cFreq], task: row[cTask], planned: cellToDateString_(row[cPlanned]), actual: cellToDateString_(row[cActual])
      });
    }
  }
  cache.put("pendingApprovals", JSON.stringify(pending), 20);
  return pending;
}

function getOverdueSummary_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("overdueSummary");
  if (cached) return JSON.parse(cached);

  const sheet = getSheet_(MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const cName = col_(map, "option", 0), cEmail = col_(map, "email", 1), cDept = col_(map, "department", 2),
        cFreq = col_(map, "freq", 4), cTask = col_(map, "task", 5), cPlanned = col_(map, "planned", 6),
        cActual = col_(map, "actual", 7);
  const cReason = map["delay reason"];

  const todayMidnight = new Date(); todayMidnight.setHours(0, 0, 0, 0);
  const lookbackStart = new Date(todayMidnight); lookbackStart.setDate(lookbackStart.getDate() - OVERDUE_LOOKBACK_DAYS);
  const overdue = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const actual = cellToDateString_(row[cActual]);
    if (actual) continue; // already done — not overdue regardless of status
    const plannedRaw = row[cPlanned];
    const plannedDate = (Object.prototype.toString.call(plannedRaw) === "[object Date]") ? plannedRaw : null;
    if (String(row[cFreq] || "").trim() === "D") continue; // Daily tasks don't carry forward as overdue
    if (!plannedDate || !(plannedDate < todayMidnight) || plannedDate < lookbackStart) continue;

    overdue.push({
      row: r + 1, name: row[cName], email: row[cEmail], department: row[cDept],
      freq: row[cFreq], task: row[cTask], planned: cellToDateString_(plannedRaw),
      delayReason: (cReason !== undefined ? (row[cReason] || "") : ""),
      _ts: plannedDate.getTime()
    });
  }
  overdue.sort((a, b) => a._ts - b._ts);
  overdue.forEach(t => delete t._ts);
  cache.put("overdueSummary", JSON.stringify(overdue), 20);
  return overdue;
}

// Admin-facing view of every task currently sitting in Rejected status,
// across all doers — pairs with the "Rejected" filter/tab on the front end.
function getRejectedSummary_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("rejectedSummary");
  if (cached) return JSON.parse(cached);

  const sheet = getSheet_(MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const cName = col_(map, "option", 0), cEmail = col_(map, "email", 1), cDept = col_(map, "department", 2),
        cFreq = col_(map, "freq", 4), cTask = col_(map, "task", 5), cPlanned = col_(map, "planned", 6),
        cStatus = col_(map, "app status", 11);
  const cRemark = firstMatchingCol_(map, REMARK_COLUMN_CANDIDATES);

  const rejected = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[cStatus] || "").trim() !== STATUS_REJECTED) continue;
    rejected.push({
      row: r + 1, name: row[cName], email: row[cEmail], department: row[cDept],
      freq: row[cFreq], task: row[cTask], planned: cellToDateString_(row[cPlanned]),
      remark: (cRemark !== undefined ? (row[cRemark] || "") : "")
    });
  }
  cache.put("rejectedSummary", JSON.stringify(rejected), 20);
  return rejected;
}

/**
 * Per-user on-time completion percentage over the last
 * PERFORMANCE_LOOKBACK_DAYS days. A task counts as "on time" if it was
 * marked done on or before its planned date; late if done after; missed
 * if still not done at all.
 */
function getUserPerformance_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("userPerformance");
  if (cached) return JSON.parse(cached);

  const sheet = getSheet_(MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const cName = col_(map, "option", 0), cEmail = col_(map, "email", 1),
        cPlanned = col_(map, "planned", 6), cActual = col_(map, "actual", 7);

  const todayMidnight = new Date(); todayMidnight.setHours(0, 0, 0, 0);
  const windowStart = new Date(todayMidnight); windowStart.setDate(windowStart.getDate() - PERFORMANCE_LOOKBACK_DAYS);

  const stats = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const email = String(row[cEmail] || "").trim().toLowerCase();
    if (!email) continue;

    const plannedRaw = row[cPlanned];
    const plannedDate = (Object.prototype.toString.call(plannedRaw) === "[object Date]") ? plannedRaw : null;
    if (!plannedDate || plannedDate > todayMidnight || plannedDate < windowStart) continue; // only due, in-window tasks count

    if (!stats[email]) stats[email] = { name: row[cName], email: row[cEmail], total: 0, onTime: 0, late: 0, missed: 0 };
    const s = stats[email];
    s.total++;

    const actualRaw = row[cActual];
    if (!actualRaw) { s.missed++; continue; }

    let actualDateOnly = null;
    if (Object.prototype.toString.call(actualRaw) === "[object Date]") {
      actualDateOnly = new Date(actualRaw.getFullYear(), actualRaw.getMonth(), actualRaw.getDate());
    } else {
      const parts = String(actualRaw).trim().split(" ")[0].split("/");
      if (parts.length === 3) actualDateOnly = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
    }
    const plannedDateOnly = new Date(plannedDate.getFullYear(), plannedDate.getMonth(), plannedDate.getDate());

    if (actualDateOnly && actualDateOnly.getTime() <= plannedDateOnly.getTime()) s.onTime++;
    else s.late++;
  }

  const result = Object.values(stats)
    .map(s => ({
      name: s.name, email: s.email, total: s.total, onTime: s.onTime, late: s.late, missed: s.missed,
      percent: s.total > 0 ? Math.round((s.onTime / s.total) * 100) : 100
    }))
    .sort((a, b) => a.percent - b.percent); // lowest performers first

  cache.put("userPerformance", JSON.stringify(result), 180); // 3 min
  return result;
}

function approveTask_(rowNumber) {
  const sheet = getSheet_(MASTER_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowNumber, col_(map, "app status", 11) + 1).setValue(STATUS_APPROVED);

  const cache = CacheService.getScriptCache();
  cache.remove("pendingApprovals");
  cache.remove("overdueSummary");
  cache.remove("rejectedSummary");

  return { row: rowNumber };
}

function rejectTask_(rowNumber, remark) {
  const sheet = getSheet_(MASTER_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const cEmail = col_(map, "email", 1);
  const doerEmail = String(sheet.getRange(rowNumber, cEmail + 1).getValue() || "").trim().toLowerCase();

  sheet.getRange(rowNumber, col_(map, "app status", 11) + 1).setValue(STATUS_REJECTED);
  sheet.getRange(rowNumber, col_(map, "actual", 7) + 1).setValue(""); // clear so doer can redo

  // Previously this only ever looked for a column literally named "remark".
  // If the Master sheet's header used a different label (e.g. "Reject
  // Reason"), cRemark came back undefined and the reason was silently
  // dropped — no error, nothing saved. Now several common header names are
  // tried, and a missing column is a loud error instead of a silent no-op.
  const cRemark = firstMatchingCol_(map, REMARK_COLUMN_CANDIDATES);
  if (cRemark === undefined) {
    throw new Error(
      'Master sheet me reject-reason ke liye koi column nahi mila. Header row me in me se ek naam add karo: "' +
      REMARK_COLUMN_CANDIDATES.join('", "') + '".'
    );
  }
  sheet.getRange(rowNumber, cRemark + 1).setValue(remark);

  const cache = CacheService.getScriptCache();
  cache.remove("pendingApprovals");
  cache.remove("overdueSummary");
  cache.remove("rejectedSummary");
  if (doerEmail) cache.remove("dash_" + doerEmail);

  return { row: rowNumber };
}

function setDelayReason_(rowNumber, reason) {
  const sheet = getSheet_(MASTER_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const cReason = map["delay reason"];
  if (cReason === undefined) {
    throw new Error('Master sheet me "Delay Reason" naam ka column nahi mila. Header row me ek column "Delay Reason" add karo.');
  }
  sheet.getRange(rowNumber, cReason + 1).setValue(reason);

  const cache = CacheService.getScriptCache();
  cache.remove("overdueSummary");

  return { row: rowNumber, reason: reason };
}

// ------------------------------------------------------------
// KEEP-WARM (optional) — reduces cold-start delay for the day's
// first request. Run setupKeepWarmTrigger() ONCE from the editor
// to schedule this; it pings the script every 10 min so it stays
// "warm" instead of cold-starting on the first real request.
// ------------------------------------------------------------
function keepWarm() {
  SpreadsheetApp.getActiveSpreadsheet().getName();
}

function setupKeepWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'keepWarm') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(10).create();
}
