// ============================================================
// smallcase B2B Growth Dashboard — Code.gs v6
// Changes from v5:
// - DevRev API integration (WhatsApp, Call Tickets, Care Emails)
// - syncDevRevNow() → incremental daily sync
// - fullBackfillDevRev() → one-time backfill from Jan 1
// - DevRev sync fires at 9:00 AM IST daily (before cache rebuilds)
// - All existing Ozonetel / cache / trigger logic unchanged
// ============================================================

// ── API KEYS ─────────────────────────────────────────────────
var CLAUDE_API_KEY = ";
var DEVREV_PAT = "";

// ── CACHE FILE ────────────────────────────────────────────────
var CACHE_FILE_NAME = "smallcase_dashboard_cache.json";

// ── SHEET NAMES ───────────────────────────────────────────────
var SHEET_NAMES = {
calls: "Ozonetel Calls",
callTkts: "Ozonetel DevRev Tickets",
whatsapp: "WhatsApp Chats",
careEmails: "Care Emails",
breaks: "Ozonetel Agent Breaks"
};

// ── COLUMN DEFINITIONS ────────────────────────────────────────
var COLS = {
calls: [
"Call Type","Call Date","Start Time","Caller No",
"Queue Time","Time to Answer","Hold Time","Talk Time","Duration",
"Wrapup Start Time","Wrapup End Time",
"Agent","Status","Dial Status","Customer Dial Status","Agent Dial Status",
"Disposition","Call Event","UCID","Call ID"
],
callTkts: [
"Title","Created date","Close date","Owner[0]","Stage",
"Metric Name[0]","Metric Name[1]","Completed In[0]","Completed In[1]",
"Metric Status[0]","Metric Status[1]",
"Grammar (5)","Maintaining SLA (15)",
"Offer further assistance & Closing statement (5)",
"Opening & Greetings (5)","Overall Score (45)",
"Escalation threat (AI)","Comments",
"Branch/Location","Broker ID (B2B)","RM Broker Name",
"Channel (B2B)","Issue Type (B2B)","RM Name","RM Number",
"Sub Issue Type (B2B)","Issue","Sub-Issue","Items"
],
whatsapp: [
"ID","Created date","Modified date","Owners[0]","Subtype",
"Branch/Location","Broker ID (B2B)","Channel (B2B)","Comments",
"Issue Type (B2B)","RM Broker Name","RM Name","RM Number",
"Sub Issue Type (B2B)","Issue","Sub-Issue",
"Metric Name[0]","Metric Name[1]","Metric Name[2]",
"Completed In[0]","Completed In[1]","Completed In[2]"
],
careEmails: [
"Title","Created date","Close date","Owner[0]","Stage",
"Reported by[0]","Account.display_name","Sentiment.label",
"Metric Name[0]","Completed In[0]","Metric Status[0]",
"Issue","Sub-Issue","Category",
"Escalation threat (AI)","Grammar (5)","Maintaining SLA (15)",
"Offer further assistance & Closing statement (5)",
"Opening & Greetings (5)","Overall Score (45)",
"Broker Name[0]","Broker ID","Items"
],
breaks: [
"Date","Agent Name",
"Break Start Time","Break End Time","Breaks","Total Break Time"
]
};

// ── Entry point — main dashboard ─────────────────────────────
function doGet(e) {
var page = e && e.parameter && e.parameter.page;
if (page === 'simple') {
return HtmlService.createHtmlOutputFromFile("simple_dashboard")
.setTitle("smallcase B2B Weekly Pulse")
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag("viewport", "width=device-width, initial-scale=1");
}
return HtmlService.createHtmlOutputFromFile("dashboard")
.setTitle("smallcase B2B Growth Dashboard")
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ================================================================
// DEVREV API — LOW-LEVEL HELPERS
// ================================================================

// Single POST call to DevRev. Throws on non-200.
function devrevPost(endpoint, body) {
var options = {
method: "post",
contentType: "application/json",
headers: { "Authorization": DEVREV_PAT },
payload: JSON.stringify(body),
muteHttpExceptions: true
};
var resp = UrlFetchApp.fetch(DEVREV_BASE + endpoint, options);
var code = resp.getResponseCode();
var text = resp.getContentText();
if (code !== 200) {
Logger.log("DevRev " + endpoint + " → HTTP " + code + ": " + text.substring(0, 800));
throw new Error("DevRev API HTTP " + code + " — " + text.substring(0, 200));
}
return JSON.parse(text);
}

// Automatically pages through all results using next_cursor.
function devrevFetchAll(endpoint, body) {
var all = [];
var cursor = null;
var maxPages = 100;

for (var page = 0; page < maxPages; page++) {
var req = JSON.parse(JSON.stringify(body));
if (cursor) req.cursor = cursor;

var resp = devrevPost(endpoint, req);
// DevRev returns works array regardless of type
var items = resp.works || [];
all = all.concat(items);

if (!resp.next_cursor || items.length === 0) break;
cursor = resp.next_cursor;
Utilities.sleep(300);
}

Logger.log("devrevFetchAll(" + endpoint + "): " + all.length + " total items");
return all;
}

// ================================================================
// DEVREV — DATE HELPER
// ================================================================
function fmtDRDate(val) {
if (!val) return "";
try {
return Utilities.formatDate(new Date(val), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
} catch(e) { return String(val); }
}

// ================================================================
// DEVREV — FIELD EXTRACTION HELPERS
// ================================================================

// Owner: DevRev puts owned_by as an array of member objects
function drOwner(r) {
var arr = r.owned_by || r.assignees || [];
if (!Array.isArray(arr)) arr = [arr];
if (!arr.length) return "";
var o = arr[0] || {};
return o.display_name || o.full_name || o.email || "";
}

// Reporter: created_by or reported_by
function drReporter(r) {
var rep = r.reported_by || r.created_by || {};
if (Array.isArray(rep)) rep = rep[0] || {};
return rep.display_name || rep.email || "";
}

// Custom fields: DevRev stores them under r.custom_fields as a flat object.
// The key names are snake_case versions of the field display names.
// We try a few variants so minor naming differences don't break things.
function drCustom(r, key) {
var cf = r.custom_fields || {};
if (cf[key] !== undefined && cf[key] !== null) return String(cf[key]);
// Try common transformations
var variants = [
key,
key.toLowerCase().replace(/\s+/g, "_"),
key.toLowerCase().replace(/[^a-z0-9]/g, "_")
];
for (var i = 0; i < variants.length; i++) {
if (cf[variants[i]] !== undefined && cf[variants[i]] !== null)
return String(cf[variants[i]]);
}
return "";
}

// SLA metric targets → [{name, value, status}]
function drMetrics(r) {
var targets = [];
if (r.sla_tracker && r.sla_tracker.metric_targets) {
targets = r.sla_tracker.metric_targets;
} else if (r.metric_targets) {
targets = r.metric_targets;
}
return targets.map(function(t) {
var name = (t.metric_definition && t.metric_definition.name) || t.name || "";
// elapsed time in minutes (how long it took)
var mins = t.completed_in_minutes != null ? t.completed_in_minutes
: t.elapsed_time_in_minutes != null ? t.elapsed_time_in_minutes
: null;
var val = mins !== null ? String(Math.round(mins * 10) / 10) : "";
var stat = t.status || (t.is_breached ? "breached" : "completed");
return { name: name, value: val, status: stat };
});
}

// ================================================================
// DEVREV — WHATSAPP CHATS
// Filter: subtype = "dealer_support" (exactly, case-insensitive)
// DevRev conversations are type "conversation" in works.list
// ================================================================
function fetchDevRevWhatsApp(sinceDate) {
// conversations.list filters: type[] accepts "conversation" or "ticket"
// created_date filter uses {after: "ISO8601"} or {$gt: "ISO8601"}
var body = {
"limit": 100
};
// DevRev date filter: use created_date.after
if (sinceDate) {
body["created_date"] = { "after": sinceDate };
}

var items = devrevFetchAll("/conversations.list", body);

// Keep only dealer_support conversations
var filtered = items.filter(function(r) {
var sub = String(r.subtype || r.conversations_subtype || "").toLowerCase().replace(/[\s\-]/g, "_");
return sub === "dealer_support";
});

Logger.log("WhatsApp dealer_support: " + filtered.length + " of " + items.length);
return filtered.map(function(r) {
var m = drMetrics(r);
return {
"ID": r.id || r.display_id || "",
"Created date": fmtDRDate(r.created_date),
"Modified date": fmtDRDate(r.modified_date),
"Owners[0]": drOwner(r),
"Subtype": r.subtype || r.conversations_subtype || "",
"Branch/Location": drCustom(r, "branch_location"),
"Broker ID (B2B)": drCustom(r, "broker_id"),
"Channel (B2B)": drCustom(r, "channel"),
"Comments": r.body || "",
"Issue Type (B2B)": drCustom(r, "issue_type"),
"RM Broker Name": drCustom(r, "rm_broker_name"),
"RM Name": drCustom(r, "rm_name"),
"RM Number": drCustom(r, "rm_number"),
"Sub Issue Type (B2B)": drCustom(r, "sub_issue_type"),
"Issue": drCustom(r, "issue"),
"Sub-Issue": drCustom(r, "sub_issue"),
"Metric Name[0]": m[0] ? m[0].name : "",
"Metric Name[1]": m[1] ? m[1].name : "",
"Metric Name[2]": m[2] ? m[2].name : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Completed In[1]": m[1] ? m[1].value : "",
"Completed In[2]": m[2] ? m[2].value : ""
};
});
}

// ================================================================
// DEVREV — CALL TICKETS
// Filter: subtype = "dealer_support"
// ================================================================
function fetchDevRevCallTickets(sinceDate) {
var body = { "type": ["ticket"], "limit": 100 };
if (sinceDate) body["created_date"] = { "after": sinceDate };

var items = devrevFetchAll("/works.list", body);

var filtered = items.filter(function(r) {
var sub = String(r.subtype || "").toLowerCase().replace(/[\s\-]/g, "_");
return sub === "dealer_support";
});

Logger.log("Call Tickets dealer_support: " + filtered.length + " of " + items.length);
return filtered.map(function(r) {
var m = drMetrics(r);
return {
"Title": r.title || "",
"Created date": fmtDRDate(r.created_date),
"Close date": fmtDRDate(r.actual_close_date || r.target_close_date || ""),
"Owner[0]": drOwner(r),
"Stage": (r.stage && r.stage.name) || "",
"Metric Name[0]": m[0] ? m[0].name : "",
"Metric Name[1]": m[1] ? m[1].name : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Completed In[1]": m[1] ? m[1].value : "",
"Metric Status[0]": m[0] ? m[0].status : "",
"Metric Status[1]": m[1] ? m[1].status : "",
"Grammar (5)": drCustom(r, "grammar"),
"Maintaining SLA (15)": drCustom(r, "maintaining_sla"),
"Offer further assistance & Closing statement (5)": drCustom(r, "closing_statement"),
"Opening & Greetings (5)": drCustom(r, "opening_greetings"),
"Overall Score (45)": drCustom(r, "overall_score"),
"Escalation threat (AI)": drCustom(r, "escalation_threat"),
"Comments": r.body || "",
"Branch/Location": drCustom(r, "branch_location"),
"Broker ID (B2B)": drCustom(r, "broker_id"),
"RM Broker Name": drCustom(r, "rm_broker_name"),
"Channel (B2B)": drCustom(r, "channel"),
"Issue Type (B2B)": drCustom(r, "issue_type"),
"RM Name": drCustom(r, "rm_name"),
"RM Number": drCustom(r, "rm_number"),
"Sub Issue Type (B2B)": drCustom(r, "sub_issue_type"),
"Issue": drCustom(r, "issue"),
"Sub-Issue": drCustom(r, "sub_issue"),
"Items": drCustom(r, "items")
};
});
}

// ================================================================
// DEVREV — CARE EMAILS
// Filter: tags include care@smallcase.com OR caresc@smallcase.com
// OR group display_name === "Care Emails"
// ================================================================
function fetchDevRevCareEmails(sinceDate) {
var body = { "type": ["ticket"], "limit": 100 };
if (sinceDate) body["created_date"] = { "after": sinceDate };

var items = devrevFetchAll("/works.list", body);

var CARE_TAGS = ["care@smallcase.com", "caresc@smallcase.com"];

var filtered = items.filter(function(r) {
// Check tags
var tags = r.tags || [];
for (var i = 0; i < tags.length; i++) {
var tn = String((tags[i] && tags[i].name) || tags[i] || "").toLowerCase().trim();
for (var j = 0; j < CARE_TAGS.length; j++) {
if (tn === CARE_TAGS[j]) return true;
}
}
// Check group
var grp = r.group || r.part || null;
if (grp) {
var gn = String(grp.display_name || grp.name || "").toLowerCase().trim();
if (gn === "care emails") return true;
}
return false;
});

Logger.log("Care Emails: " + filtered.length + " of " + items.length);
return filtered.map(function(r) {
var m = drMetrics(r);
return {
"Title": r.title || "",
"Created date": fmtDRDate(r.created_date),
"Close date": fmtDRDate(r.actual_close_date || r.target_close_date || ""),
"Owner[0]": drOwner(r),
"Stage": (r.stage && r.stage.name) || "",
"Reported by[0]": drReporter(r),
"Account.display_name": (r.rev_org && r.rev_org.display_name) || "",
"Sentiment.label": (r.sentiment && r.sentiment.label) || "",
"Metric Name[0]": m[0] ? m[0].name : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Metric Status[0]": m[0] ? m[0].status : "",
"Issue": drCustom(r, "issue"),
"Sub-Issue": drCustom(r, "sub_issue"),
"Category": drCustom(r, "category"),
"Escalation threat (AI)": drCustom(r, "escalation_threat"),
"Grammar (5)": drCustom(r, "grammar"),
"Maintaining SLA (15)": drCustom(r, "maintaining_sla"),
"Offer further assistance & Closing statement (5)": drCustom(r, "closing_statement"),
"Opening & Greetings (5)": drCustom(r, "opening_greetings"),
"Overall Score (45)": drCustom(r, "overall_score"),
"Broker Name[0]": drCustom(r, "rm_broker_name"),
"Broker ID": drCustom(r, "broker_id"),
"Items": drCustom(r, "items")
};
});
}

// ================================================================
// DEVREV — WRITE TO SHEET (upsert / append-new-only)
// ================================================================
function writeDevRevToSheet(ss, sheetName, newRows, colNames) {
if (!newRows || newRows.length === 0) {
Logger.log(sheetName + ": 0 new rows — nothing to write");
return 0;
}

var sheet = ss.getSheetByName(sheetName);
if (!sheet) {
sheet = ss.insertSheet(sheetName);
Logger.log("Created sheet: " + sheetName);
}

var data = sheet.getDataRange().getValues();

// Write header row if sheet is empty
if (data.length === 0) {
sheet.appendRow(colNames);
data = [colNames];
}

var headers = data[0].map(function(h) { return String(h).trim(); });

// Ensure all required columns exist as headers
var headersChanged = false;
colNames.forEach(function(col) {
if (headers.indexOf(col) === -1) {
headers.push(col);
headersChanged = true;
}
});
if (headersChanged) {
sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// Build deduplication key set from existing data
// Key = ID (conversations) or Title (tickets) + Created date
var idIdx = headers.indexOf("ID");
var titleIdx = headers.indexOf("Title");
var dateIdx = headers.indexOf("Created date");
var existingKeys = {};
for (var i = 1; i < data.length; i++) {
var pk = idIdx >= 0 ? String(data[i][idIdx]).trim()
: titleIdx >= 0 ? String(data[i][titleIdx]).trim() : "";
var dk = dateIdx >= 0 ? String(data[i][dateIdx]).trim() : "";
if (pk) existingKeys[pk + "||" + dk] = true;
}

// Append only records not already present
var added = 0;
newRows.forEach(function(row) {
var pk = (row["ID"] || row["Title"] || "").trim();
var dk = (row["Created date"] || "").trim();
if (pk && existingKeys[pk + "||" + dk]) return; // skip duplicate

var rowArr = headers.map(function(col) {
return row[col] !== undefined ? row[col] : "";
});
sheet.appendRow(rowArr);
added++;
});

Logger.log(sheetName + ": added " + added + " new rows (skipped " + (newRows.length - added) + " duplicates)");
return added;
}

// ================================================================
// DEVREV — HOURLY SYNC (Mon–Fri, 9:05 AM – 7:05 PM IST)
// Also fills gaps from previous working day if needed
// ================================================================
function syncDevRevHourly() {
var now = new Date();
// Convert to IST (UTC+5:30)
var istOffset = 5.5 * 60 * 60000;
var istNow = new Date(now.getTime() + istOffset);
var dayOfWeek = istNow.getUTCDay(); // 0=Sun, 1=Mon ... 5=Fri, 6=Sat
var hour = istNow.getUTCHours(); // 0-23 in IST

// Only run Mon–Fri (1–5) between 9:00 AM and 7:59 PM IST
if (dayOfWeek === 0 || dayOfWeek === 6) {
Logger.log("syncDevRevHourly: skipping — weekend");
return;
}
if (hour < 9 || hour >= 20) {
Logger.log("syncDevRevHourly: skipping — outside business hours (IST hour=" + hour + ")");
return;
}

Logger.log("syncDevRevHourly: running at IST hour=" + hour);
syncDevRevNow();
}

// ================================================================
// DEVREV — MASTER SYNC (incremental — only fetches since last run)
// Also checks if previous working day has a gap and fills it
// ================================================================
function syncDevRevNow() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var props = PropertiesService.getScriptProperties();

// Read last sync timestamp; default to start of 2026
var lastSync = props.getProperty("devrev_last_sync") || "2026-01-01T00:00:00Z";

// Gap detection: if last sync was > 14 hours ago, go back further to catch any missed window
try {
var lastSyncDate = new Date(lastSync);
var hoursSince = (Date.now() - lastSyncDate.getTime()) / 3600000;
if (hoursSince > 14) {
// Find start of the previous working day (Mon–Fri) as the safe "since" date
var safeBack = new Date(lastSyncDate);
safeBack.setHours(0, 0, 0, 0);
lastSync = Utilities.formatDate(safeBack, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
Logger.log("Gap detected (" + Math.round(hoursSince) + "h) — extending since to: " + lastSync);
}
} catch(e) {
Logger.log("Gap check error (non-fatal): " + e.message);
}

Logger.log("DevRev sync since: " + lastSync);

try {
var waRows = fetchDevRevWhatsApp(lastSync);
var ctRows = fetchDevRevCallTickets(lastSync);
var emRows = fetchDevRevCareEmails(lastSync);

writeDevRevToSheet(ss, SHEET_NAMES.whatsapp, waRows, COLS.whatsapp);
writeDevRevToSheet(ss, SHEET_NAMES.callTkts, ctRows, COLS.callTkts);
writeDevRevToSheet(ss, SHEET_NAMES.careEmails, emRows, COLS.careEmails);

var now = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
props.setProperty("devrev_last_sync", now);
Logger.log("DevRev sync complete at: " + now);

buildDashboardCache();

return {
success: true, syncedAt: now,
fetched: { whatsapp: waRows.length, callTickets: ctRows.length, careEmails: emRows.length }
};

} catch(e) {
Logger.log("DevRev sync FAILED: " + e.message);
throw e;
}
}

// ================================================================
// DEVREV — FULL BACKFILL (run once to load all historical data)
// Resets the last-sync timestamp and fetches everything from Jan 1 2026
// ================================================================
function fullBackfillDevRev() {
var props = PropertiesService.getScriptProperties();
props.deleteProperty("devrev_last_sync");
var result = syncDevRevNow();
var msg = "Fetched from Jan 1 2026 to now:\n" +
" WhatsApp chats: " + result.fetched.whatsapp + "\n" +
" Call Tickets: " + result.fetched.callTickets + "\n" +
" Care Emails: " + result.fetched.careEmails + "\n\n" +
"Cache rebuilt. Reload the dashboard to see data.";
Logger.log("Full Backfill complete — " + msg);
try { SpreadsheetApp.getUi().alert("Full Backfill Complete", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

// ================================================================
// DEVREV — DEBUG: inspect one raw record of each type
// Run this from the Apps Script editor, then check View → Logs
// to verify field names and fix any drCustom() key mismatches
// ================================================================
function inspectDevRevFields() {
Logger.log("=== INSPECTING ONE TICKET (call tickets / care emails) ===");
try {
var tResp = devrevPost("/works.list", { type: ["ticket"], limit: 1 });
var ticket = (tResp.works || [])[0];
if (ticket) {
Logger.log("Top-level keys: " + Object.keys(ticket).join(", "));
Logger.log("custom_fields: " + JSON.stringify(ticket.custom_fields || {}, null, 2));
Logger.log("sla_tracker: " + JSON.stringify(ticket.sla_tracker || {}, null, 2));
Logger.log("tags: " + JSON.stringify(ticket.tags || [], null, 2));
Logger.log("stage: " + JSON.stringify(ticket.stage || {}, null, 2));
Logger.log("owned_by: " + JSON.stringify(ticket.owned_by || [], null, 2));
Logger.log("group: " + JSON.stringify(ticket.group || {}, null, 2));
Logger.log("subtype: " + (ticket.subtype || "(none)"));
} else {
Logger.log("No tickets found.");
}
} catch(e) { Logger.log("Ticket inspect error: " + e.message); }

Logger.log("=== INSPECTING ONE CONVERSATION (WhatsApp) ===");
try {
var cResp = devrevPost("/works.list", { type: ["conversation"], limit: 1 });
var conv = (cResp.works || [])[0];
if (conv) {
Logger.log("Top-level keys: " + Object.keys(conv).join(", "));
Logger.log("custom_fields: " + JSON.stringify(conv.custom_fields || {}, null, 2));
Logger.log("subtype: " + (conv.subtype || conv.conversations_subtype || "(none)"));
Logger.log("owned_by: " + JSON.stringify(conv.owned_by || [], null, 2));
Logger.log("sla_tracker: " + JSON.stringify(conv.sla_tracker || {}, null, 2));
} else {
Logger.log("No conversations found.");
}
} catch(e) { Logger.log("Conversation inspect error: " + e.message); }

Logger.log("=== DONE — check View > Logs ===");
}

// ================================================================
// SHEET READER (Ozonetel — unchanged from v5)
// ================================================================
function readSlimSheet(sheetName, neededCols) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetName);
if (!sheet) { Logger.log("Sheet not found: " + sheetName); return []; }

var data = sheet.getDataRange().getValues();
if (data.length < 2) return [];

var headers = data[0].map(function(h) { return String(h).trim(); });
var colMap = {};
neededCols.forEach(function(col) {
var idx = headers.indexOf(col);
if (idx !== -1) colMap[col] = idx;
});

var rows = [];
for (var i = 1; i < data.length; i++) {
var row = data[i];
var hasData = false;
for (var j = 0; j < row.length; j++) {
if (row[j] !== "" && row[j] !== null && row[j] !== undefined) { hasData = true; break; }
}
if (!hasData) continue;

var obj = {};
neededCols.forEach(function(col) {
var idx = colMap[col];
if (idx === undefined) { obj[col] = ""; return; }
var val = row[idx];
if (val instanceof Date) {
obj[col] = Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
} else if (val === null || val === undefined) {
obj[col] = "";
} else {
obj[col] = val;
}
});
rows.push(obj);
}
return rows;
}

// ================================================================
// DRIVE CACHE
// ================================================================
function getCacheFile() {
var files = DriveApp.getFilesByName(CACHE_FILE_NAME);
if (files.hasNext()) return files.next();
return DriveApp.createFile(CACHE_FILE_NAME, "{}", MimeType.PLAIN_TEXT);
}

function buildDashboardCache() {
Logger.log("Building dashboard cache...");
var start = Date.now();

var payload = {
calls: readSlimSheet(SHEET_NAMES.calls, COLS.calls),
callTkts: readSlimSheet(SHEET_NAMES.callTkts, COLS.callTkts),
whatsapp: readSlimSheet(SHEET_NAMES.whatsapp, COLS.whatsapp),
careEmails: readSlimSheet(SHEET_NAMES.careEmails, COLS.careEmails),
breaks: readSlimSheet(SHEET_NAMES.breaks, COLS.breaks),
builtAt: Utilities.formatDate(new Date(), "Asia/Kolkata", "dd MMM yyyy, hh:mm a") + " IST",
rowCounts: {}
};
payload.rowCounts.calls = payload.calls.length;
payload.rowCounts.callTkts = payload.callTkts.length;
payload.rowCounts.whatsapp = payload.whatsapp.length;
payload.rowCounts.careEmails = payload.careEmails.length;
payload.rowCounts.breaks = payload.breaks.length;

getCacheFile().setContent(JSON.stringify(payload));

var elapsed = ((Date.now() - start) / 1000).toFixed(1);
Logger.log("Cache built in " + elapsed + "s — " +
payload.rowCounts.calls + " calls, " + payload.rowCounts.callTkts + " tickets, " +
payload.rowCounts.whatsapp + " WA, " + payload.rowCounts.careEmails + " emails");

return { success: true, builtAt: payload.builtAt, rowCounts: payload.rowCounts, elapsed: elapsed };
}

function getAllData() {
try {
var file = getCacheFile();
var content = file.getBlob().getDataAsString();
if (!content || content === "{}") { Logger.log("Cache empty — reading sheets directly"); return buildAndReturn(); }
var data = JSON.parse(content);
if (!data.calls || !data.callTkts) { Logger.log("Cache malformed — rebuilding"); return buildAndReturn(); }
Logger.log("Serving from cache (built: " + (data.builtAt || "unknown") + ")");
return data;
} catch(e) {
Logger.log("Cache read failed: " + e.message + " — falling back to sheets");
return buildAndReturn();
}
}

function buildAndReturn() {
buildDashboardCache();
return JSON.parse(getCacheFile().getBlob().getDataAsString());
}

function forceRebuildCache() { return buildDashboardCache(); }

// ================================================================
// TRIGGER SETUP — run setupAllTriggers() ONCE after deployment
// ================================================================
function setupAllTriggers() {
// Remove old triggers to avoid duplicates
var existing = ScriptApp.getProjectTriggers();
var deleted = 0;
existing.forEach(function(t) {
var fn = t.getHandlerFunction();
if (["buildDashboardCache","onChangeHandler","autoRebuildAfterPaste","syncDevRevNow","syncDevRevHourly"].indexOf(fn) !== -1) {
ScriptApp.deleteTrigger(t); deleted++;
}
});
if (deleted > 0) Logger.log("Removed " + deleted + " old trigger(s)");

// 1. DevRev hourly sync Mon–Fri 9:05 AM – 7:05 PM IST
// Apps Script time-based triggers fire every N hours; we use everyHours(1).
// The syncDevRevHourly() function checks day-of-week and hour before running.
ScriptApp.newTrigger("syncDevRevHourly").timeBased().everyHours(1).create();
Logger.log("Hourly DevRev sync trigger created");

// 2. Catch-up sync at 9:05 AM every day (also handles previous-day gap fill)
ScriptApp.newTrigger("syncDevRevNow").timeBased().everyDays(1).atHour(9).nearMinute(5).create();

// 3. Cache-only rebuilds as safety net
[{ h:10, m:30 }, { h:11, m:30 }].forEach(function(s) {
ScriptApp.newTrigger("buildDashboardCache").timeBased().everyDays(1).atHour(s.h).nearMinute(s.m).create();
});

// 4. onChange — auto-detect Ozonetel CSV paste
ScriptApp.newTrigger("onChangeHandler")
.forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
.onChange()
.create();

Logger.log("All triggers created successfully. Check View > Logs for confirmation.");

// Try to show UI alert — works when run from sheet menu, silently skips when run from editor
try {
SpreadsheetApp.getUi().alert(
"All Triggers Created ✓",
"Active schedule (Asia/Kolkata):\n\n" +
" Every hour (Mon–Fri, 9:05–19:05) — DevRev sync\n" +
" 9:05 AM daily — DevRev catch-up (fills previous day gaps)\n" +
" 10:30 AM & 11:30 AM — Cache rebuild\n" +
" On Ozonetel paste (~15s) — Auto date-fix + cache rebuild\n\n" +
"Next: Open Support Dashboard menu → 📥 Full Backfill from Jan 1 2026",
SpreadsheetApp.getUi().ButtonSet.OK
);
} catch(e) {
Logger.log("Triggers created. (UI alert skipped — run from Sheet menu for the popup)");
}
}

function setupDailyTriggers() { setupAllTriggers(); }
function setupDailyTrigger() { setupAllTriggers(); }

function listTriggers() {
var triggers = ScriptApp.getProjectTriggers();
var msg = triggers.length === 0
? "No triggers set up."
: triggers.map(function(t) { return t.getHandlerFunction() + " — " + t.getEventType(); }).join("\n");
Logger.log("Active triggers: " + msg);
try { SpreadsheetApp.getUi().alert("Active Triggers", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

// ================================================================
// CUSTOM MENU
// ================================================================
function onOpen() {
SpreadsheetApp.getUi()
.createMenu("Support Dashboard")
.addItem("⚡ Setup All Triggers (run once)", "setupAllTriggers")
.addItem("Check Active Triggers", "listTriggers")
.addSeparator()
.addItem("🔄 Sync DevRev Now (incremental)", "syncDevRevNow")
.addItem("📥 Full Backfill from Jan 1 2026", "fullBackfillDevRev")
.addItem("🔍 Inspect DevRev Fields (debug)", "inspectDevRevFields")
.addSeparator()
.addItem("Force Rebuild Cache Now", "forceRebuildCacheFromMenu")
.addItem("Fix All Dates (manual fallback)", "fixAllDates")
.addSeparator()
.addItem("Open Dashboard (Advanced)", "openDashboard")
.addItem("Open Weekly Pulse (Simple)", "openSimpleDashboard")
.addItem("Test: Row Counts", "testRowCounts")
.addToUi();
}

function openSimpleDashboard() {
var url = ScriptApp.getService().getUrl() + '?page=simple';
var html = HtmlService.createHtmlOutput(
'<script>window.open("' + url + '"); google.script.host.close();</script>'
).setWidth(10).setHeight(10);
try { SpreadsheetApp.getUi().showModalDialog(html, "Opening Weekly Pulse…"); } catch(e) {}
}

function forceRebuildCacheFromMenu() {
var result = buildDashboardCache();
var msg = "Built at: " + result.builtAt + "\n\n" +
"Rows loaded:\n" +
" Calls: " + result.rowCounts.calls + "\n" +
" Call Tickets: " + result.rowCounts.callTkts + "\n" +
" WhatsApp: " + result.rowCounts.whatsapp + "\n" +
" Care Emails: " + result.rowCounts.careEmails + "\n\n" +
"Built in " + result.elapsed + "s.\nReload the dashboard to see new data.";
Logger.log(msg);
try { SpreadsheetApp.getUi().alert("Cache Rebuilt", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

function openDashboard() {
var url = ScriptApp.getService().getUrl();
var html = HtmlService.createHtmlOutput(
'<script>window.open("' + url + '"); google.script.host.close();</script>'
).setWidth(10).setHeight(10);
SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
}

function testRowCounts() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var results = [];
Object.keys(SHEET_NAMES).forEach(function(key) {
var name = SHEET_NAMES[key];
var sheet = ss.getSheetByName(name);
if (!sheet) { results.push(name + ": NOT FOUND"); return; }
var slim = readSlimSheet(name, COLS[key]);
results.push(name + ": " + slim.length + " rows");
});
try {
var file = getCacheFile();
var content = file.getBlob().getDataAsString();
var cached = JSON.parse(content);
results.push("");
results.push("Cache: " + CACHE_FILE_NAME);
results.push("Built: " + (cached.builtAt || "never"));
results.push("Size: " + Math.round(content.length / 1024) + " KB");
var props = PropertiesService.getScriptProperties();
var lastSync = props.getProperty("devrev_last_sync") || "(never synced)";
results.push("Last DevRev sync: " + lastSync);
} catch(e) { results.push("Cache: not built yet"); }
Logger.log(results.join("\n"));
try { SpreadsheetApp.getUi().alert("Data Status", results.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

// ================================================================
// AI SUMMARY
// ================================================================
function getAISummary(comments, dateLabel) {
if (!CLAUDE_API_KEY) return { error: "API key not set in Code.gs." };
if (!comments || comments.length === 0) return { error: "No comments found for the selected date range." };

var safeComments = comments.slice(0, 200).map(function(c, i) {
return (i + 1) + ". " + String(c).replace(/[\r\n\t]+/g, " ").replace(/[^\x20-\x7E]/g, "").replace(/"/g, "'").trim().substring(0, 200);
}).join("\n");

var prompt =
"You are analysing B2B support ticket comments from smallcase (a fintech platform) for: " + dateLabel + ".\n\n" +
"COMMENTS:\n" + safeComments + "\n\n" +
"TASK: Group these comments into 5-8 themes. Then write a 3-sentence narrative summary.\n\n" +
"RULES:\n- Respond ONLY with raw JSON. No markdown. No code fences.\n" +
"- Use only ASCII characters.\n" +
'- Structure: {"clusters":[{"theme":"Payment Links","count":12,"examples":["ex1","ex2"]}],"narrative":"Summary."}';

try {
var resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
method: "post", contentType: "application/json",
headers: { "x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01" },
payload: JSON.stringify({ model: "claude-haiku-4-5-20251001", max_tokens: 1500, messages: [{ role: "user", content: prompt }] }),
muteHttpExceptions: true
});
var result = JSON.parse(resp.getContentText());
if (result.error) return { error: result.error.message };
var text = result.content[0].text.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
var jsonStr = extractOutermostJSON(text);
if (!jsonStr) return { error: "Could not parse response. Raw: " + text.substring(0, 200) };
var parsed = JSON.parse(jsonStr);
if (!parsed.clusters) parsed.clusters = [];
if (!parsed.narrative) parsed.narrative = "Summary not available.";
return parsed;
} catch(e) { return { error: "API error: " + e.message }; }
}

function extractOutermostJSON(text) {
var start = text.indexOf("{");
if (start === -1) return null;
var depth = 0, inString = false, escape = false;
for (var i = start; i < text.length; i++) {
var ch = text[i];
if (escape) { escape = false; continue; }
if (ch === "\\" && inString) { escape = true; continue; }
if (ch === '"') { inString = !inString; continue; }
if (inString) continue;
if (ch === "{") depth++;
else if (ch === "}") { depth--; if (depth === 0) return text.substring(start, i + 1); }
}
return null;
}

// ================================================================
// DATE FIX (Ozonetel dates — unchanged)
// ================================================================
var DATE_COLUMNS = {
"Ozonetel Calls": ["Call Date"],
"Ozonetel DevRev Tickets": ["Created date","Close date","Modified date"],
"WhatsApp Chats": ["Created date","Modified date"],
"Care Emails": ["Created date","Close date","Modified date"],
"Ozonetel Agent Breaks": ["Date"]
};

function fixAllDates() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var totalFixed = 0;
Object.keys(DATE_COLUMNS).forEach(function(sheetName) {
var sheet = ss.getSheetByName(sheetName);
if (!sheet) return;
var data = sheet.getDataRange().getValues();
if (data.length < 2) return;
var headers = data[0].map(function(h) { return String(h).trim(); });
DATE_COLUMNS[sheetName].forEach(function(colName) {
var colIdx = headers.indexOf(colName);
if (colIdx === -1) return;
var colLetter = columnToLetter(colIdx + 1);
sheet.getRange(colLetter + "2:" + colLetter + data.length).setNumberFormat("@");
for (var row = 1; row < data.length; row++) {
var cell = data[row][colIdx];
if (!cell || cell === "") continue;
var parsed = normaliseDateCell(cell);
if (parsed) {
sheet.getRange(row + 1, colIdx + 1).setValue(Utilities.formatDate(parsed, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss"));
totalFixed++;
}
}
});
});
buildDashboardCache();
var msg = "Fixed " + totalFixed + " date cells and rebuilt cache.";
Logger.log(msg);
try { SpreadsheetApp.getUi().alert("Done", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

function normaliseDateCell(cell) {
if (cell instanceof Date && !isNaN(cell.getTime())) return cell;
var s = String(cell).trim();
if (!s) return null;
var m1 = s.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})/);
if (m1) {
var day = parseInt(m1[1]), mon = parseInt(m1[2]), yr = parseInt(m1[3]);
if (yr < 100) yr += 2000;
if (day >= 1 && day <= 31 && mon >= 1 && mon <= 12) {
var tp = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
var h=0,mi=0,sec=0;
if (tp) { h=parseInt(tp[1]); mi=parseInt(tp[2]); sec=parseInt(tp[3]||0); }
return new Date(yr, mon-1, day, h, mi, sec);
}
}
var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
if (m2) {
var tp2 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
var h2=0,mi2=0,sec2=0;
if (tp2) { h2=parseInt(tp2[1]); mi2=parseInt(tp2[2]); sec2=parseInt(tp2[3]||0); }
return new Date(parseInt(m2[1]), parseInt(m2[2])-1, parseInt(m2[3]), h2, mi2, sec2);
}
var d = new Date(s);
return isNaN(d.getTime()) ? null : d;
}

function columnToLetter(col) {
var letter = "";
while (col > 0) { var rem = (col-1)%26; letter = String.fromCharCode(65+rem)+letter; col = Math.floor((col-1)/26); }
return letter;
}

// ================================================================
// AUTO PASTE DETECTION (Ozonetel)
// ================================================================
var REBUILD_LOCK_KEY = "rebuildQueued";

function onChangeHandler(e) {
try {
if (e && e.changeType !== "INSERT_ROWS" && e.changeType !== "OTHER") return;
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetName = ss.getActiveSheet().getName();
var watched = [SHEET_NAMES.calls, SHEET_NAMES.callTkts, SHEET_NAMES.whatsapp, SHEET_NAMES.careEmails, SHEET_NAMES.breaks];
if (watched.indexOf(sheetName) === -1) return;
var cache = CacheService.getScriptCache();
if (cache.get(REBUILD_LOCK_KEY)) { Logger.log("Rebuild already queued — skipping"); return; }
cache.put(REBUILD_LOCK_KEY, "1", 60);
Logger.log("Paste detected in " + sheetName + " — scheduling rebuild in 15s");
ScriptApp.newTrigger("autoRebuildAfterPaste").timeBased().after(15 * 1000).create();
} catch(err) { Logger.log("onChangeHandler error: " + err.message); }
}

function autoRebuildAfterPaste() {
try {
Logger.log("autoRebuildAfterPaste: starting...");
CacheService.getScriptCache().remove(REBUILD_LOCK_KEY);
fixAllDatesSilent();
var result = buildDashboardCache();
Logger.log("Auto-rebuild complete: " + result.builtAt);
} catch(err) {
Logger.log("autoRebuildAfterPaste error: " + err.message);
} finally {
ScriptApp.getProjectTriggers().forEach(function(t) {
if (t.getHandlerFunction() === "autoRebuildAfterPaste") ScriptApp.deleteTrigger(t);
});
}
}

function fixAllDatesSilent() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var totalFixed = 0;
Object.keys(DATE_COLUMNS).forEach(function(sheetName) {
var sheet = ss.getSheetByName(sheetName);
if (!sheet) return;
var data = sheet.getDataRange().getValues();
if (data.length < 2) return;
var headers = data[0].map(function(h) { return String(h).trim(); });
DATE_COLUMNS[sheetName].forEach(function(colName) {
var colIdx = headers.indexOf(colName);
if (colIdx === -1) return;
var colLetter = columnToLetter(colIdx + 1);
sheet.getRange(colLetter + "2:" + colLetter + data.length).setNumberFormat("@");
for (var row = 1; row < data.length; row++) {
var cell = data[row][colIdx];
if (!cell || cell === "") continue;
var parsed = normaliseDateCell(cell);
if (parsed) { sheet.getRange(row+1, colIdx+1).setValue(Utilities.formatDate(parsed, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss")); totalFixed++; }
}
});
});
Logger.log("fixAllDatesSilent: fixed " + totalFixed + " date cells");
}
