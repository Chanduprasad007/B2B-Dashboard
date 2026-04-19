// ============================================================
// smallcase B2B Growth Dashboard — Code.gs v7
// DevRev:
// - Tickets: api.devrev.ai /works.list
// - WhatsApp: app.devrev.ai internal gateway /conversations.list + created_date (public api.devrev.ai rejects created_date here)
// - syncDevRevYesterdayIST() — daily import of items created *yesterday* (IST window)
// - syncDevRevNow() — incremental since last run (manual / optional)
// Time-based triggers use the Apps Script project time zone — set to Asia/Kolkata for 9:00 IST.
// ============================================================

// ── API KEYS ─────────────────────────────────────────────────
var CLAUDE_API_KEY = "";
var DEVREV_PAT     = "";
var DEVREV_BASE = "https://api.devrev.ai";
var DEVREV_INTERNAL_CONVERSATIONS_LIST = "https://app.devrev.ai/api/gateway/internal/conversations.list";

// ── CACHE FILE ────────────────────────────────────────────────
var CACHE_FILE_NAME = "smallcase_dashboard_cache.json";
var BACKFILL_PROGRESS_KEY = "devrev_backfill_last_completed_date";

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
"Disposition","Call Event","UCID","Call ID","Recording URL"
],
callTkts: ["Work ID"].concat(
"Title,Close date,Created by,Created date,Modified by,Modified date,Owner[0],Reported by[0],Stage,Tags[0],Tags[1],Tags[2],Tags[3],Tags[4],Tags[5],Work type,Part,Group,Immutable,Severity.color,Severity.id,Severity.label,Severity.ordinal,SLA Name.display_id,SLA Name.id,SLA Name.name,SLA Name.status,Metric Name[0],Metric Name[1],Metric Stage[0],Metric Stage[1],Completed In[0],Completed In[1],Metric Status[0],Metric Status[1],Subtype,Visibility.id,Visibility.label,Visibility.ordinal,Acknowledgement and Assurance (15),Agent Messages,Broker ID,Broker Name[0],Clickup link,Create Clickup,Customer Messages,Escalation threat (AI),Grammar (5),Last Agent Message Timestamp,Maintaining SLA (15),Offer further assistance & Closing statement (5),Opening & Greetings (5),Overall Score (45),SAM ID,smallboard URL,Branch/Location,Broker ID (B2B),RM Broker Name,Channel (B2B),Comments,Issue Type (B2B),RM Name,RM Number,Sub Issue Type (B2B),Recording URL,Items".split(",")
).concat(["Record JSON"]),
whatsapp: [
"ID","Last Message","Modified date","Created date","Owners[0]","Created by","Subtype",
"Stage",
"Branch/Location","Broker ID (B2B)","Channel (B2B)","Comments",
"Issue Type (B2B)","RM Broker Name","RM Name","RM Number",
"Sub Issue Type (B2B)",
"Metric Name[0]","Metric Name[1]","Metric Name[2]",
"Metric Stage[0]","Metric Stage[1]","Metric Stage[2]",
"Completed In[0]","Completed In[1]","Completed In[2]",
"Record JSON"
],
careEmails: ["Work ID"].concat(
"Title,Close date,Body,Created by,Created date,Links[0],Modified by,Modified date,Owner[0],Reported by[0],Reported by[1],Reported by[2],Reported by[3],Reported by[4],Reported by[5],Reported by[6],Reported by[7],Reported by[8],Reported by[9],Stage,Tags[0],Tags[1],Tags[2],Tags[3],Tags[4],Work type,Account.display_id,Account.display_name,Account.id,Account.id_v1,Account.is_archived,Account.thumbnail,Part,Channels[0],Group,Immutable,Workspace,Sentiment.id,Sentiment.label,Sentiment.ordinal,Severity.color,Severity.id,Severity.label,Severity.ordinal,SLA Name.display_id,SLA Name.id,SLA Name.name,SLA Name.status,Metric Name[0],Metric Name[1],Metric Stage[0],Metric Stage[1],Completed In[0],Completed In[1],Metric Status[0],Metric Status[1],Source channel,CSAT Rating[0].average,CSAT Rating[0].count,CSAT Rating[0].maximum,CSAT Rating[0].minimum,CSAT Rating[0].sum,CSAT Rating[0].survey_id,CSAT Rating[0].survey_id_v1,Visibility.id,Visibility.label,Visibility.ordinal,Acknowledgement and Assurance (15),Agent Messages,Broker ID,Category,Clickup link,Create Clickup,Customer Messages,Escalation threat (AI),Grammar (5),Issue,Last Agent Message Timestamp,Maintaining SLA (15),Offer further assistance & Closing statement (5),Opening & Greetings (5),Overall Score (45),SAM ID,smallboard URL,Sub-Issue,Items".split(",")
).concat(["Record JSON"]),
breaks: [
"Date","Agent Name",
"Break Start Time","Break End Time","Breaks","Total Break Time"
]
};

var CACHE_COLS = {
calls: COLS.calls.slice(),
callTkts: COLS.callTkts.filter(function(col) { return col !== "Record JSON"; }),
whatsapp: COLS.whatsapp.filter(function(col) { return col !== "Record JSON"; }),
careEmails: COLS.careEmails.filter(function(col) { return col !== "Record JSON"; }),
breaks: COLS.breaks.slice()
};

// ── Entry point — main dashboard ─────────────────────────────
function doGet(e) {
var page = e && e.parameter && e.parameter.page;
if (page === 'simple') {
return HtmlService.createHtmlOutputFromFile("simple_dashboard")
.setTitle("smallcase B2B Weekly Pulse")
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag("viewport", "width=device-width, initial-scale=1")
.setFaviconUrl("https://i.ibb.co/zWjPBDY1/sc-simple.png")
;
}
return HtmlService.createHtmlOutputFromFile("dashboard")
.setTitle("smallcase B2B Growth Dashboard")
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag("viewport", "width=device-width, initial-scale=1")
.setFaviconUrl("https://i.ibb.co/kgmD6R7s/sc-advanced.png");
}

// ================================================================
// DEVREV API — LOW-LEVEL HELPERS
// ================================================================

// DevRev expects Authorization: Bearer <PAT>
function devrevAuthHeader() {
var t = String(DEVREV_PAT || "").trim();
if (!t) return "";
if (t.toLowerCase().indexOf("bearer ") === 0) return t;
return "Bearer " + t;
}

// Single POST to a full URL (internal gateway or other).
function devrevPostFullUrl(url, body) {
var options = {
method: "post",
contentType: "application/json",
headers: {
"Authorization": devrevAuthHeader(),
"Accept": "application/json"
},
payload: JSON.stringify(body),
muteHttpExceptions: true
};
var resp = UrlFetchApp.fetch(url, options);
var code = resp.getResponseCode();
var text = resp.getContentText();
if (code !== 200) {
Logger.log("DevRev request POST body: " + JSON.stringify(body));
Logger.log("DevRev " + url + " → HTTP " + code + ": " + text.substring(0, 2000));
throw new Error("DevRev API HTTP " + code + " — " + text.substring(0, 200));
}
return JSON.parse(text);
}

// Single POST call to public API (api.devrev.ai). Throws on non-200.
function devrevPost(endpoint, body) {
return devrevPostFullUrl(DEVREV_BASE + endpoint, body);
}

// Automatically pages through all results using next_cursor.
// itemsKey: "works" (works.list) or "conversations" (conversations.list)
function devrevFetchAll(endpoint, body, itemsKey) {
return devrevFetchAllAtUrl(DEVREV_BASE + endpoint, body, itemsKey);
}

function devrevFetchAllAtUrl(fullUrl, body, itemsKey) {
var key = itemsKey || "works";
var all = [];
var cursor = null;
var maxPages = 100;

for (var page = 0; page < maxPages; page++) {
var req = JSON.parse(JSON.stringify(body));
if (cursor) req.cursor = cursor;

var resp = devrevPostFullUrl(fullUrl, req);
var items = resp[key] || [];
all = all.concat(items);

if (!resp.next_cursor || items.length === 0) break;
cursor = resp.next_cursor;
Utilities.sleep(300);
}

Logger.log("devrevFetchAllAtUrl(" + fullUrl + ", " + key + "): " + all.length + " total items");
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

// ISO range for "created yesterday" in Asia/Kolkata (after = start of yesterday IST, before = start of today IST).
function getYesterdayISTCreatedRange() {
var tz = "Asia/Kolkata";
var now = new Date();
var todayYmd = Utilities.formatDate(now, tz, "yyyy-MM-dd");
var parts = todayYmd.split("-");
var y = parseInt(parts[0], 10), mo = parseInt(parts[1], 10) - 1, da = parseInt(parts[2], 10);
var pad = function(n) { return (n < 10 ? "0" : "") + n; };
var isoStartToday = y + "-" + pad(mo + 1) + "-" + pad(da) + "T00:00:00+05:30";
var startToday = new Date(isoStartToday);
var startYesterday = new Date(startToday.getTime() - 24 * 60 * 60 * 1000);
return { after: startYesterday.toISOString(), before: startToday.toISOString() };
}

function getYesterdayISTModifiedRange() {
return getYesterdayISTCreatedRange();
}

function getISTRangeForTodayWindow(startHour, startMinute, endHourExclusive, endMinuteExclusive) {
var tz = "Asia/Kolkata";
var now = new Date();
var todayYmd = Utilities.formatDate(now, tz, "yyyy-MM-dd");
var parts = todayYmd.split("-");
var y = parseInt(parts[0], 10), mo = parseInt(parts[1], 10) - 1, da = parseInt(parts[2], 10);
var pad = function(n) { return (n < 10 ? "0" : "") + n; };
var isoStart = y + "-" + pad(mo + 1) + "-" + pad(da) + "T" + pad(startHour) + ":" + pad(startMinute) + ":00+05:30";
var isoEnd = y + "-" + pad(mo + 1) + "-" + pad(da) + "T" + pad(endHourExclusive) + ":" + pad(endMinuteExclusive) + ":00+05:30";
return { after: new Date(isoStart).toISOString(), before: new Date(isoEnd).toISOString(), label: todayYmd + " " + pad(startHour) + ":" + pad(startMinute) + "–" + pad(endHourExclusive) + ":" + pad(endMinuteExclusive) + " IST" };
}

// One calendar day in IST: created_date in [day 00:00, next day 00:00) — dateStr = "yyyy-MM-dd" (IST date).
function getISTCreatedRangeForCalendarDay(dateStr) {
var parts = String(dateStr).trim().split("-");
if (parts.length !== 3) throw new Error("getISTCreatedRangeForCalendarDay: use yyyy-MM-dd, got " + dateStr);
var y = parseInt(parts[0], 10), mo = parseInt(parts[1], 10) - 1, da = parseInt(parts[2], 10);
var pad = function(n) { return (n < 10 ? "0" : "") + n; };
var isoStart = y + "-" + pad(mo + 1) + "-" + pad(da) + "T00:00:00+05:30";
var startDay = new Date(isoStart);
var endNext = new Date(startDay.getTime() + 24 * 60 * 60 * 1000);
return { after: startDay.toISOString(), before: endNext.toISOString(), label: parts[0] + "-" + pad(parseInt(parts[1], 10)) + "-" + pad(da) };
}

// Generic created_date (e.g. internal gateway conversations).
function applyCreatedDateFilter(body, sinceIso, untilIso) {
if (!sinceIso && !untilIso) return;
var f = {};
if (sinceIso) f.after = sinceIso;
if (untilIso) f.before = untilIso;
body.created_date = f;
}

function applyModifiedDateFilter(body, sinceIso, untilIso) {
if (!sinceIso && !untilIso) return;
var f = {};
if (sinceIso) f.after = sinceIso;
if (untilIso) f.before = untilIso;
body.modified_date = f;
}

// works.list tickets — created_date must use type: "range" and inclusive end-of-window before.
// Example: { type: "range", after: "2026-04-11T18:30:00.000Z", before: "2026-04-12T18:29:59.999Z" }
// untilIso = exclusive start of next period (e.g. midnight IST today); before = untilIso − 1 ms.
function applyWorksListCreatedDateFilter(body, sinceIso, untilIso) {
if (!sinceIso && !untilIso) return;
var f = { type: "range" };
if (sinceIso) f.after = sinceIso;
if (untilIso) {
f.before = new Date(new Date(untilIso).getTime() - 1).toISOString();
}
body.created_date = f;
}

function applyWorksListModifiedDateFilter(body, sinceIso, untilIso) {
if (!sinceIso && !untilIso) return;
var f = { type: "range" };
if (sinceIso) f.after = sinceIso;
if (untilIso) {
f.before = new Date(new Date(untilIso).getTime() - 1).toISOString();
}
body.modified_date = f;
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

// Try multiple custom_fields keys (tenant-prefixed tnt__/ctype__/app_ plus short names).
function drCf(r, keys) {
var cf = r.custom_fields || {};
if (!keys || !keys.length) return "";
for (var i = 0; i < keys.length; i++) {
var k = keys[i];
if (k === undefined || k === null || k === "") continue;
var v = cf[k];
if (v !== undefined && v !== null && v !== "") return String(v);
}
var last = keys[keys.length - 1];
return last ? drCustom(r, last) : "";
}

function drSeverity(r) {
var s = r.severity;
if (s == null || s === "") return { id: "", label: "", color: "", ordinal: "" };
if (typeof s === "string") return { id: "", label: s, color: "", ordinal: "" };
if (typeof s === "object") {
return {
id: s.id != null ? String(s.id) : "",
label: s.label || "",
color: s.color || "",
ordinal: s.ordinal != null ? String(s.ordinal) : ""
};
}
return { id: "", label: "", color: "", ordinal: "" };
}

// SLA metric targets → [{name, value, status, stage}] (works.list + sla_summary shapes)
function drMetrics(r) {
var targets = [];
var st = r.sla_tracker;
if (st && st.metric_targets && st.metric_targets.length) {
targets = st.metric_targets;
} else if (r.metric_targets && r.metric_targets.length) {
targets = r.metric_targets;
} else if (r.sla_summary && r.sla_summary.sla_tracker && r.sla_summary.sla_tracker.metric_target_summaries) {
targets = r.sla_summary.sla_tracker.metric_target_summaries;
} else if (r.sla_summary && r.sla_summary.metric_target_summaries && r.sla_summary.metric_target_summaries.length) {
targets = r.sla_summary.metric_target_summaries;
}
return targets.map(function(t) {
var name = (t.metric_definition && t.metric_definition.name) || t.name || "";
var mins = t.completed_in_minutes != null ? t.completed_in_minutes
: t.elapsed_time_in_minutes != null ? t.elapsed_time_in_minutes
: t.completed_in != null ? t.completed_in
: null;
var val = mins !== null ? String(Math.round(mins * 10) / 10) : "";
var stat = t.status || (t.is_breached ? "breached" : "");
var stg = t.stage || (t.metric_stage && t.metric_stage.name) || "";
return { name: name, value: val, status: stat, stage: stg };
});
}

// Last message reference string (matches DevRev export "Last Message" column when present).
function drLastMessageRef(r) {
var msgs = r.messages;
if (!msgs || !msgs.length) return "";
var last = msgs[msgs.length - 1];
if (!last) return "";
return last.object || last.object_display_id || last.external_ref || last.id || "";
}

function conversationBody(r) {
return r.body || r.description || "";
}

function safeJsonStringify(obj) {
try {
var s = JSON.stringify(obj);
return s.length > 48000 ? s.substring(0, 48000) + "…" : s;
} catch (e) {
return "";
}
}

// Same string for sheet cell vs incoming row so dedupe works (Sheets stores dates as Date objects).
function normalizeDedupeCell(val) {
if (val === null || val === undefined) return "";
if (val instanceof Date && !isNaN(val.getTime())) {
return Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}
return String(val).trim();
}

function normalizeStageName(val) {
return String(val || "").trim().toLowerCase();
}

function isClosedLikeStage(val) {
var s = normalizeStageName(val);
return s === "resolved" || s === "archived" || s === "canceled" || s === "cancelled";
}

function parseRecordJsonId(val) {
var s = normalizeDedupeCell(val);
if (!s) return "";
try {
var obj = JSON.parse(s);
return String(obj.id || obj.display_id || "").trim();
} catch (e) {
return "";
}
}

function rowArrayToObject(headers, rowArr) {
var obj = {};
for (var i = 0; i < headers.length; i++) obj[headers[i]] = rowArr[i] !== undefined ? rowArr[i] : "";
return obj;
}

function devrevRowTimestamp(row) {
var cols = ["Modified date", "Close date", "Created date"];
for (var i = 0; i < cols.length; i++) {
var d = normaliseDateCell(row[cols[i]]);
if (d && !isNaN(d.getTime())) return d.getTime();
}
return 0;
}

function devrevRowScore(row) {
var score = 0;
Object.keys(row || {}).forEach(function(k) {
if (row[k] !== "" && row[k] !== null && row[k] !== undefined) score++;
});
return score;
}

function choosePreferredDevRevRow(existingRow, newRow) {
if (!existingRow) return newRow;
if (!newRow) return existingRow;
var oldTs = devrevRowTimestamp(existingRow);
var newTs = devrevRowTimestamp(newRow);
if (newTs > oldTs) return newRow;
if (oldTs > newTs) return existingRow;
return devrevRowScore(newRow) >= devrevRowScore(existingRow) ? newRow : existingRow;
}

function devrevStableKey(sheetName, row) {
var workId = String(row["Work ID"] || "").trim();
var displayId = String(row["ID"] || "").trim();
var rawId = parseRecordJsonId(row["Record JSON"]);
if (sheetName === SHEET_NAMES.whatsapp) {
if (displayId) return "wa:id:" + displayId;
if (rawId) return "wa:raw:" + rawId;
}
if (workId) return sheetName + ":work:" + workId;
if (displayId) return sheetName + ":id:" + displayId;
if (rawId) return sheetName + ":raw:" + rawId;
var title = String(row["Title"] || row["Last Message"] || "").trim();
var created = normalizeDedupeCell(row["Created date"]);
if (title || created) return sheetName + ":fallback:" + title + "||" + created;
return "";
}

function drCreatedBy(r) {
var c = r.created_by || {};
if (Array.isArray(c)) c = c[0] || {};
return c.display_name || c.display_id || c.email || "";
}

function drModifiedBy(r) {
var c = r.modified_by || {};
if (Array.isArray(c)) c = c[0] || {};
return c.display_name || c.display_id || c.email || "";
}

function drTagAt(r, i) {
var tags = r.tags || [];
var t = tags[i];
if (!t) return "";
return ((t.tag && t.tag.name) || t.name || t.value || "").toString();
}

function drReportedByAt(r, i) {
var arr = r.reported_by || [];
var x = arr[i];
if (!x) return "";
return x.display_name || x.display_id || x.email || "";
}

function drWorkDisplayId(r) {
return r.display_id || "";
}

function mapCallTicketExportRow(r) {
var m = drMetrics(r);
var sev = drSeverity(r);
var vis = r.visibility;
if (vis != null && typeof vis !== "object") vis = { id: vis, label: "", ordinal: "" };
if (!vis) vis = {};
var slaTr = r.sla_tracker || {};
var sla = slaTr.sla || r.sla || {};
var part = r.applies_to_part || r.part || {};
var grp = r.group || r.part || {};
var imm = r.immutable;
return {
"Work ID": drWorkDisplayId(r),
"Title": r.title || "",
"Close date": fmtDRDate(r.actual_close_date || r.target_close_date || ""),
"Created by": drCreatedBy(r),
"Created date": fmtDRDate(r.created_date),
"Modified by": drModifiedBy(r),
"Modified date": fmtDRDate(r.modified_date),
"Owner[0]": drOwner(r),
"Reported by[0]": drReportedByAt(r, 0),
"Stage": (r.stage && r.stage.name) || "",
"Tags[0]": drTagAt(r, 0),
"Tags[1]": drTagAt(r, 1),
"Tags[2]": drTagAt(r, 2),
"Tags[3]": drTagAt(r, 3),
"Tags[4]": drTagAt(r, 4),
"Tags[5]": drTagAt(r, 5),
"Work type": r.type || "ticket",
"Part": part.name || "",
"Group": grp.display_name || grp.name || "",
"Immutable": imm === true ? "true" : imm === false ? "false" : "",
"Severity.color": sev.color || "",
"Severity.id": sev.id != null ? String(sev.id) : "",
"Severity.label": sev.label || "",
"Severity.ordinal": sev.ordinal != null ? String(sev.ordinal) : "",
"SLA Name.display_id": sla.display_id || slaTr.sla_id || "",
"SLA Name.id": (sla.id && String(sla.id)) || "",
"SLA Name.name": sla.name || slaTr.sla_name || "",
"SLA Name.status": sla.status || "",
"Metric Name[0]": m[0] ? m[0].name : "",
"Metric Name[1]": m[1] ? m[1].name : "",
"Metric Stage[0]": m[0] ? m[0].stage : "",
"Metric Stage[1]": m[1] ? m[1].stage : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Completed In[1]": m[1] ? m[1].value : "",
"Metric Status[0]": m[0] ? m[0].status : "",
"Metric Status[1]": m[1] ? m[1].status : "",
"Subtype": r.subtype || "",
"Visibility.id": vis.id != null ? String(vis.id) : "",
"Visibility.label": vis.label || "",
"Visibility.ordinal": vis.ordinal != null ? String(vis.ordinal) : "",
"Acknowledgement and Assurance (15)": drCf(r, ["tnt__acknolwedgement_and_assurance_15", "tnt__acknowledgement_and_assurance_15", "acknowledgement_and_assurance"]),
"Agent Messages": drCf(r, ["tnt__agent_messages", "agent_messages"]),
"Broker ID": drCf(r, ["tnt__broker_id", "broker_id"]),
"Broker Name[0]": drCf(r, ["ctype__broker_name", "tnt__broker_name", "rm_broker_name"]),
"Clickup link": drCf(r, ["tnt__clickup_link", "clickup_link", "tnt__create_clickup"]),
"Create Clickup": drCf(r, ["tnt__create_clickup", "create_clickup"]),
"Customer Messages": drCf(r, ["tnt__customer_messages", "customer_messages"]),
"Escalation threat (AI)": drCf(r, ["tnt__escalation_threat_ai", "escalation_threat"]),
"Grammar (5)": drCf(r, ["tnt__grammar_5", "grammar"]),
"Last Agent Message Timestamp": fmtDRDate(drCf(r, ["tnt__last_agent_message_timestamp", "last_agent_message_timestamp"]) || r.last_agent_message_timestamp || r.last_agent_message_date),
"Maintaining SLA (15)": drCf(r, ["tnt__maintaining_sla_15", "maintaining_sla"]),
"Offer further assistance & Closing statement (5)": drCf(r, ["tnt__offer_further_assitance_closing_statement_5", "tnt__offer_further_assistance_closing_statement_5", "closing_statement"]),
"Opening & Greetings (5)": drCf(r, ["tnt__opening_greetings_5", "opening_greetings"]),
"Overall Score (45)": drCf(r, ["tnt__overall_score_100", "tnt__overall_score_45", "overall_score"]),
"SAM ID": drCf(r, ["tnt__sam_id", "sam_id"]),
"smallboard URL": drCf(r, ["tnt__smallboard_url", "smallboard_url"]),
"Branch/Location": drCf(r, ["ctype__branch_location", "branch_location"]),
"Broker ID (B2B)": drCf(r, ["ctype__broker_id", "tnt__broker_id_b2b", "broker_id_b2b", "broker_id"]),
"RM Broker Name": drCf(r, ["ctype__rm_broker_name", "tnt__rm_broker_name", "ctype__broker_name", "tnt__broker_name", "rm_broker_name"]),
"Channel (B2B)": drCf(r, ["ctype__channel", "channel"]),
"Comments": drCf(r, ["ctype__comments", "comments"]) || (r.body || ""),
"Issue Type (B2B)": drCf(r, ["ctype__issue_type_b2b", "issue_type"]),
"RM Name": drCf(r, ["ctype__rm_name", "rm_name"]),
"RM Number": drCf(r, ["ctype__rm_number", "rm_number"]),
"Sub Issue Type (B2B)": drCf(r, ["ctype__sub_issue_type_b2b", "sub_issue_type"]),
"Recording URL": drCf(r, ["recording_url", "ctype__recording_url", "ozonetel_link", "call_recording_url"]),
"Items": drCf(r, ["tnt__items", "items"]),
"Record JSON": safeJsonStringify(r)
};
}

function mapCareEmailExportRow(r) {
var m = drMetrics(r);
var sev = drSeverity(r);
var vis = r.visibility;
if (vis != null && typeof vis !== "object") vis = { id: vis, label: "", ordinal: "" };
if (!vis) vis = {};
var slaTr = r.sla_tracker || {};
var sla = slaTr.sla || r.sla || {};
var acc = r.account || r.rev_org || {};
var grp = r.group || r.part || {};
var sent = r.sentiment || {};
var csat = (r.csat_rating && r.csat_rating[0]) || r.csat_rating || {};
return {
"Work ID": drWorkDisplayId(r),
"Title": r.title || "",
"Close date": fmtDRDate(r.actual_close_date || r.target_close_date || ""),
"Body": r.body || "",
"Created by": drCreatedBy(r),
"Created date": fmtDRDate(r.created_date),
"Links[0]": r.external_ref || "",
"Modified by": drModifiedBy(r),
"Modified date": fmtDRDate(r.modified_date),
"Owner[0]": drOwner(r),
"Reported by[0]": drReportedByAt(r, 0),
"Reported by[1]": drReportedByAt(r, 1),
"Reported by[2]": drReportedByAt(r, 2),
"Reported by[3]": drReportedByAt(r, 3),
"Reported by[4]": drReportedByAt(r, 4),
"Reported by[5]": drReportedByAt(r, 5),
"Reported by[6]": drReportedByAt(r, 6),
"Reported by[7]": drReportedByAt(r, 7),
"Reported by[8]": drReportedByAt(r, 8),
"Reported by[9]": drReportedByAt(r, 9),
"Stage": (r.stage && r.stage.name) || "",
"Tags[0]": drTagAt(r, 0),
"Tags[1]": drTagAt(r, 1),
"Tags[2]": drTagAt(r, 2),
"Tags[3]": drTagAt(r, 3),
"Tags[4]": drTagAt(r, 4),
"Work type": r.type || "ticket",
"Account.display_id": acc.display_id || "",
"Account.display_name": acc.display_name || "",
"Account.id": acc.id || "",
"Account.id_v1": acc.id_v1 || "",
"Account.is_archived": acc.is_archived === true ? "true" : acc.is_archived === false ? "false" : "",
"Account.thumbnail": acc.thumbnail || "",
"Part": (r.applies_to_part && r.applies_to_part.name) || "",
"Channels[0]": (function() {
var ch = r.channels;
if (!ch || !ch.length) return "";
var c0 = ch[0];
return typeof c0 === "string" ? c0 : (c0 && c0.name) || "";
})(),
"Group": grp.display_name || grp.name || "",
"Immutable": r.immutable === true ? "true" : r.immutable === false ? "false" : "",
"Workspace": drCustom(r, "workspace"),
"Sentiment.id": sent.id != null ? String(sent.id) : "",
"Sentiment.label": sent.label || "",
"Sentiment.ordinal": sent.ordinal != null ? String(sent.ordinal) : "",
"Severity.color": sev.color || "",
"Severity.id": sev.id != null ? String(sev.id) : "",
"Severity.label": sev.label || "",
"Severity.ordinal": sev.ordinal != null ? String(sev.ordinal) : "",
"SLA Name.display_id": sla.display_id || "",
"SLA Name.id": (sla.id && String(sla.id)) || "",
"SLA Name.name": sla.name || "",
"SLA Name.status": sla.status || "",
"Metric Name[0]": m[0] ? m[0].name : "",
"Metric Name[1]": m[1] ? m[1].name : "",
"Metric Stage[0]": m[0] ? m[0].stage : "",
"Metric Stage[1]": m[1] ? m[1].stage : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Completed In[1]": m[1] ? m[1].value : "",
"Metric Status[0]": m[0] ? m[0].status : "",
"Metric Status[1]": m[1] ? m[1].status : "",
"Source channel": r.source_channel || "",
"CSAT Rating[0].average": csat.average != null ? String(csat.average) : "",
"CSAT Rating[0].count": csat.count != null ? String(csat.count) : "",
"CSAT Rating[0].maximum": csat.maximum != null ? String(csat.maximum) : "",
"CSAT Rating[0].minimum": csat.minimum != null ? String(csat.minimum) : "",
"CSAT Rating[0].sum": csat.sum != null ? String(csat.sum) : "",
"CSAT Rating[0].survey_id": csat.survey_id || "",
"CSAT Rating[0].survey_id_v1": csat.survey_id_v1 || "",
"Visibility.id": vis.id != null ? String(vis.id) : "",
"Visibility.label": vis.label || "",
"Visibility.ordinal": vis.ordinal != null ? String(vis.ordinal) : "",
"Acknowledgement and Assurance (15)": drCf(r, ["tnt__acknolwedgement_and_assurance_15", "tnt__acknowledgement_and_assurance_15", "acknowledgement_and_assurance"]),
"Agent Messages": drCf(r, ["tnt__agent_messages", "agent_messages"]),
"Broker ID": drCf(r, ["tnt__broker_id", "broker_id"]),
"Category": drCf(r, ["tnt__category", "category"]),
"Clickup link": drCf(r, ["tnt__clickup_link", "clickup_link", "tnt__create_clickup"]),
"Create Clickup": drCf(r, ["tnt__create_clickup", "create_clickup"]),
"Customer Messages": drCf(r, ["tnt__customer_messages", "customer_messages"]),
"Escalation threat (AI)": drCf(r, ["tnt__escalation_threat_ai", "escalation_threat"]),
"Grammar (5)": drCf(r, ["tnt__grammar_5", "grammar"]),
"Issue": drCf(r, ["tnt__issue", "issue"]),
"Last Agent Message Timestamp": fmtDRDate(drCf(r, ["tnt__last_agent_message_timestamp", "last_agent_message_timestamp"]) || r.last_agent_message_timestamp || r.last_agent_message_date),
"Maintaining SLA (15)": drCf(r, ["tnt__maintaining_sla_15", "maintaining_sla"]),
"Offer further assistance & Closing statement (5)": drCf(r, ["tnt__offer_further_assitance_closing_statement_5", "tnt__offer_further_assistance_closing_statement_5", "closing_statement"]),
"Opening & Greetings (5)": drCf(r, ["tnt__opening_greetings_5", "opening_greetings"]),
"Overall Score (45)": drCf(r, ["tnt__overall_score_100", "tnt__overall_score_45", "overall_score"]),
"SAM ID": drCf(r, ["tnt__sam_id", "sam_id"]),
"smallboard URL": drCf(r, ["tnt__smallboard_url", "smallboard_url"]),
"Sub-Issue": drCf(r, ["tnt__sub_issue", "sub_issue"]),
"Items": drCf(r, ["tnt__items", "items"]),
"Record JSON": safeJsonStringify(r)
};
}

// ================================================================
// DEVREV — WHATSAPP CHATS (internal gateway conversations.list + created_date)
// Public api.devrev.ai/conversations.list rejects created_date; app gateway accepts it.
// ================================================================
function fetchDevRevWhatsApp(sinceIso, untilIso, opts) {
opts = opts || {};
var dateField = opts.dateField || "created";
var onlyOpen = opts.onlyOpen === true;
var body = { "limit": 100, "subtype": ["dealer_support"] };
if (dateField === "modified") applyModifiedDateFilter(body, sinceIso, untilIso);
else applyCreatedDateFilter(body, sinceIso, untilIso);

var items = devrevFetchAllAtUrl(DEVREV_INTERNAL_CONVERSATIONS_LIST, body, "conversations");

var rangeStart = sinceIso ? new Date(sinceIso).getTime() : null;
var rangeEnd = untilIso ? new Date(untilIso).getTime() : null;

var filtered = items.filter(function(r) {
var sub = String(r.subtype || r.conversations_subtype || "").toLowerCase().replace(/[\s\-]/g, "_");
if (sub !== "dealer_support") return false;
var stageName = normalizeStageName((r.stage && (r.stage.name || r.stage.display_name || r.stage.id)) || r.stage || "");
if (onlyOpen && (!stageName || isClosedLikeStage(stageName))) return false;
var rawDate = dateField === "modified" ? r.modified_date : r.created_date;
if (!rawDate) return false;
var ct = new Date(rawDate).getTime();
if (rangeStart !== null && ct < rangeStart) return false;
if (rangeEnd !== null && ct >= rangeEnd) return false;
return true;
});

Logger.log("WhatsApp (conversations.list) dealer_support " + dateField + "_date" + (onlyOpen ? " open-only" : "") + ": " + filtered.length + " of " + items.length);
return filtered.map(function(r) {
var m = drMetrics(r);
return {
"ID": r.display_id || r.id || "",
"Last Message": drLastMessageRef(r),
"Modified date": fmtDRDate(r.modified_date),
"Created date": fmtDRDate(r.created_date),
"Owners[0]": drOwner(r),
"Created by": drReporter(r),
"Subtype": r.subtype || "",
"Stage": (r.stage && r.stage.name) || "",
"Branch/Location": drCf(r, ["ctype__branch_location", "branch_location"]),
"Broker ID (B2B)": drCf(r, ["ctype__broker_id_b2b", "ctype__broker_id", "broker_id"]),
"Channel (B2B)": drCf(r, ["ctype__channel_b2b", "ctype__channel", "channel"]),
"Comments": drCf(r, ["ctype__comments", "comments"]) || conversationBody(r),
"Issue Type (B2B)": drCf(r, ["ctype__issue_type_b2b", "issue_type"]),
"RM Broker Name": drCf(r, ["ctype__rm_broker_name", "tnt__rm_broker_name", "ctype__broker_name", "tnt__broker_name", "rm_broker_name"]),
"RM Name": drCf(r, ["ctype__rm_name", "rm_name"]),
"RM Number": drCf(r, ["ctype__rm_number", "rm_number"]),
"Sub Issue Type (B2B)": drCf(r, ["ctype__sub_issue_type_b2b", "sub_issue_type"]),
"Metric Name[0]": m[0] ? m[0].name : "",
"Metric Name[1]": m[1] ? m[1].name : "",
"Metric Name[2]": m[2] ? m[2].name : "",
"Metric Stage[0]": m[0] ? m[0].stage : "",
"Metric Stage[1]": m[1] ? m[1].stage : "",
"Metric Stage[2]": m[2] ? m[2].stage : "",
"Completed In[0]": m[0] ? m[0].value : "",
"Completed In[1]": m[1] ? m[1].value : "",
"Completed In[2]": m[2] ? m[2].value : "",
"Record JSON": safeJsonStringify(r)
};
});
}

// ================================================================
// DEVREV — CALL TICKETS
// API filter: ticket.subtype + subtype_op (works.list)
// ================================================================
function fetchDevRevCallTickets(sinceIso, untilIso, opts) {
opts = opts || {};
var dateField = opts.dateField || "created";
var onlyOpen = opts.onlyOpen === true;
var body = {
"type": ["ticket"],
"limit": 100,
"ticket": {
"subtype": ["dealer_support"]
}
};
if (dateField === "modified") applyWorksListModifiedDateFilter(body, sinceIso, untilIso);
else applyWorksListCreatedDateFilter(body, sinceIso, untilIso);

var items = devrevFetchAll("/works.list", body, "works");

var rangeStart = sinceIso ? new Date(sinceIso).getTime() : null;
var rangeEnd = untilIso ? new Date(untilIso).getTime() : null;

var filtered = items.filter(function(r) {
var sub = String(r.subtype || "").toLowerCase().replace(/[\s\-]/g, "_");
if (sub !== "dealer_support") return false;
var stageName = normalizeStageName((r.stage && (r.stage.name || r.stage.display_name || r.stage.id)) || r.stage || "");
if (onlyOpen && (!stageName || isClosedLikeStage(stageName))) return false;
var rawDate = dateField === "modified" ? r.modified_date : r.created_date;
if (rangeStart !== null && rangeEnd !== null && rawDate) {
var ct = new Date(rawDate).getTime();
if (ct < rangeStart || ct >= rangeEnd) return false;
}
return true;
});

Logger.log("Call Tickets dealer_support " + dateField + "_date" + (onlyOpen ? " open-only" : "") + ": " + filtered.length + " of " + items.length);
return filtered.map(function(r) { return mapCallTicketExportRow(r); });
}

// ================================================================
// DEVREV — CARE EMAILS
// Filter: tags include care@smallcase.com OR caresc@smallcase.com
// OR group display_name === "Care Emails"
// ================================================================
function fetchDevRevCareEmails(sinceIso, untilIso, opts) {
opts = opts || {};
var dateField = opts.dateField || "created";
var onlyOpen = opts.onlyOpen === true;
var body = { "type": ["ticket"], "limit": 100 };
if (dateField === "modified") applyWorksListModifiedDateFilter(body, sinceIso, untilIso);
else applyWorksListCreatedDateFilter(body, sinceIso, untilIso);

var items = devrevFetchAll("/works.list", body, "works");

var CARE_TAGS = ["care@smallcase.com", "caresc@smallcase.com"];

var rangeStart = sinceIso ? new Date(sinceIso).getTime() : null;
var rangeEnd = untilIso ? new Date(untilIso).getTime() : null;

var filtered = items.filter(function(r) {
// Check tags
var tags = r.tags || [];
var care = false;
for (var i = 0; i < tags.length; i++) {
var tn = String((tags[i] && tags[i].tag && tags[i].tag.name) || (tags[i] && tags[i].name) || tags[i] || "").toLowerCase().trim();
for (var j = 0; j < CARE_TAGS.length; j++) {
if (tn === CARE_TAGS[j]) { care = true; break; }
}
if (care) break;
}
if (!care) {
var grp = r.group || r.part || null;
if (grp) {
var gn = String(grp.display_name || grp.name || "").toLowerCase().trim();
if (gn === "care emails") care = true;
}
}
if (!care) return false;
var stageName = normalizeStageName((r.stage && (r.stage.name || r.stage.display_name || r.stage.id)) || r.stage || "");
if (onlyOpen && (!stageName || isClosedLikeStage(stageName))) return false;
var rawDate = dateField === "modified" ? r.modified_date : r.created_date;
if (rangeStart !== null && rangeEnd !== null && rawDate) {
var ct = new Date(rawDate).getTime();
if (ct < rangeStart || ct >= rangeEnd) return false;
}
return true;
});

Logger.log("Care Emails " + dateField + "_date" + (onlyOpen ? " open-only" : "") + ": " + filtered.length + " of " + items.length);
return filtered.map(function(r) { return mapCareEmailExportRow(r); });
}

// ================================================================
// DEVREV — WRITE TO SHEET (upsert / append-new-only)
// ================================================================
function writeDevRevToSheet(ss, sheetName, newRows, colNames) {
newRows = newRows || [];

var sheet = ss.getSheetByName(sheetName);
if (!sheet) {
sheet = ss.insertSheet(sheetName);
Logger.log("Created sheet: " + sheetName);
}

var data = sheet.getDataRange().getValues();

// Write header row if sheet is empty (also when API returns 0 rows — new sheet stays usable)
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
data = sheet.getDataRange().getValues();
headers = data[0].map(function(h) { return String(h).trim(); });
}

if (newRows.length === 0) {
Logger.log(sheetName + ": 0 new data rows — tab ready with headers");
return 0;
}

var merged = {};
var existingOrder = [];

for (var i = 1; i < data.length; i++) {
var existingObj = rowArrayToObject(headers, data[i]);
var existingKey = devrevStableKey(sheetName, existingObj);
if (!existingKey) continue;
if (!merged[existingKey]) existingOrder.push(existingKey);
merged[existingKey] = choosePreferredDevRevRow(merged[existingKey], existingObj);
}

var incoming = {};
newRows.forEach(function(row) {
var key = devrevStableKey(sheetName, row);
if (!key) return;
incoming[key] = choosePreferredDevRevRow(incoming[key], row);
});

var added = 0, updated = 0;
Object.keys(incoming).forEach(function(key) {
if (merged[key]) updated++;
else { added++; existingOrder.push(key); }
merged[key] = choosePreferredDevRevRow(merged[key], incoming[key]);
});

existingOrder = existingOrder.filter(function(key, idx, arr) { return merged[key] && arr.indexOf(key) === idx; });
 existingOrder.sort(function(a, b) {
 var aCreated = normaliseDateCell(merged[a]["Created date"]);
 var bCreated = normaliseDateCell(merged[b]["Created date"]);
 var aCreatedTs = aCreated && !isNaN(aCreated.getTime()) ? aCreated.getTime() : 0;
 var bCreatedTs = bCreated && !isNaN(bCreated.getTime()) ? bCreated.getTime() : 0;
 if (aCreatedTs !== bCreatedTs) return aCreatedTs - bCreatedTs;
 return devrevRowTimestamp(merged[a]) - devrevRowTimestamp(merged[b]);
 });

var outputRows = existingOrder.map(function(key) {
var row = merged[key];
return headers.map(function(col) { return row[col] !== undefined ? row[col] : ""; });
});

sheet.clearContents();
sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
if (outputRows.length) {
sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
}

Logger.log(sheetName + ": upserted " + Object.keys(incoming).length + " unique rows — added " + added + ", updated " + updated + ", final sheet rows " + outputRows.length);
return added;
}

function appendDevRevRowsToSheet(ss, sheetName, newRows, colNames) {
newRows = newRows || [];

var sheet = ss.getSheetByName(sheetName);
if (!sheet) {
sheet = ss.insertSheet(sheetName);
Logger.log("Created sheet: " + sheetName);
}

var data = sheet.getDataRange().getValues();
if (data.length === 0) {
sheet.getRange(1, 1, 1, colNames.length).setValues([colNames]);
data = [colNames];
}

var headers = data[0].map(function(h) { return String(h).trim(); });
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

if (newRows.length === 0) {
Logger.log(sheetName + ": 0 backfill rows for this day");
return 0;
}

var incoming = {};
Object.keys(newRows).forEach(function(idx) {
var row = newRows[idx];
var key = devrevStableKey(sheetName, row);
if (!key) return;
incoming[key] = choosePreferredDevRevRow(incoming[key], row);
});

var orderedKeys = Object.keys(incoming);
orderedKeys.sort(function(a, b) {
var aCreated = normaliseDateCell(incoming[a]["Created date"]);
var bCreated = normaliseDateCell(incoming[b]["Created date"]);
var aCreatedTs = aCreated && !isNaN(aCreated.getTime()) ? aCreated.getTime() : 0;
var bCreatedTs = bCreated && !isNaN(bCreated.getTime()) ? bCreated.getTime() : 0;
if (aCreatedTs !== bCreatedTs) return aCreatedTs - bCreatedTs;
return devrevRowTimestamp(incoming[a]) - devrevRowTimestamp(incoming[b]);
});

var outputRows = orderedKeys.map(function(key) {
var row = incoming[key];
return headers.map(function(col) { return row[col] !== undefined ? row[col] : ""; });
});

if (outputRows.length) {
var startRow = Math.max(sheet.getLastRow() + 1, 2);
sheet.getRange(startRow, 1, outputRows.length, headers.length).setValues(outputRows);
}

Logger.log(sheetName + ": appended " + outputRows.length + " unique backfill rows");
return outputRows.length;
}

function minIso(a, b) {
if (!a) return b;
if (!b) return a;
return new Date(a).getTime() <= new Date(b).getTime() ? a : b;
}

function getEarliestUnresolvedCreatedIso(ss, sheetName) {
var sheet = ss.getSheetByName(sheetName);
if (!sheet) return null;
var data = sheet.getDataRange().getValues();
if (data.length < 2) return null;
var headers = data[0].map(function(h) { return String(h).trim(); });
var stageIdx = headers.indexOf("Stage");
var createdIdx = headers.indexOf("Created date");
if (createdIdx === -1) return null;

var earliest = null;
for (var i = 1; i < data.length; i++) {
var stage = stageIdx >= 0 ? normalizeStageName(data[i][stageIdx]) : "";
if (!stage) continue;
if (isClosedLikeStage(stage)) continue;
var created = normaliseDateCell(data[i][createdIdx]);
if (!created || isNaN(created.getTime())) continue;
if (!earliest || created.getTime() < earliest.getTime()) earliest = created;
}
return earliest ? earliest.toISOString() : null;
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
var waSince = minIso(lastSync, getEarliestUnresolvedCreatedIso(ss, SHEET_NAMES.whatsapp));
var ctSince = minIso(lastSync, getEarliestUnresolvedCreatedIso(ss, SHEET_NAMES.callTkts));
var emSince = minIso(lastSync, getEarliestUnresolvedCreatedIso(ss, SHEET_NAMES.careEmails));

Logger.log("WhatsApp sync window since: " + waSince);
Logger.log("Call Tickets sync window since: " + ctSince);
Logger.log("Care Emails sync window since: " + emSince);

var waRows = fetchDevRevWhatsApp(waSince, null);
var ctRows = fetchDevRevCallTickets(ctSince, null);
var emRows = fetchDevRevCareEmails(emSince, null);

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
// DEVREV — DAILY: items created YESTERDAY (IST) — for 9:00 AM trigger
// ================================================================
function syncDevRevYesterdayIST() {
var createdRange = getYesterdayISTCreatedRange();
var modifiedRange = getYesterdayISTModifiedRange();
Logger.log("syncDevRevYesterdayIST created_date after=" + createdRange.after + " before=" + createdRange.before);
Logger.log("syncDevRevYesterdayIST modified_date after=" + modifiedRange.after + " before=" + modifiedRange.before + " (open-only)");

var ss = SpreadsheetApp.getActiveSpreadsheet();
try {
var waCreated = fetchDevRevWhatsApp(createdRange.after, createdRange.before);
var ctCreated = fetchDevRevCallTickets(createdRange.after, createdRange.before);
var emCreated = fetchDevRevCareEmails(createdRange.after, createdRange.before);

var waModifiedOpen = fetchDevRevWhatsApp(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });
var ctModifiedOpen = fetchDevRevCallTickets(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });
var emModifiedOpen = fetchDevRevCareEmails(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });

var waRows = waCreated.concat(waModifiedOpen);
var ctRows = ctCreated.concat(ctModifiedOpen);
var emRows = emCreated.concat(emModifiedOpen);

writeDevRevToSheet(ss, SHEET_NAMES.whatsapp, waRows, COLS.whatsapp);
writeDevRevToSheet(ss, SHEET_NAMES.callTkts, ctRows, COLS.callTkts);
writeDevRevToSheet(ss, SHEET_NAMES.careEmails, emRows, COLS.careEmails);

var props = PropertiesService.getScriptProperties();
var stamp = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
props.setProperty("devrev_last_yesterday_sync", stamp);
props.setProperty("devrev_last_yesterday_label", Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd"));

buildDashboardCache();

Logger.log("syncDevRevYesterdayIST complete — WA created:" + waCreated.length + " modified-open:" + waModifiedOpen.length + " | CT created:" + ctCreated.length + " modified-open:" + ctModifiedOpen.length + " | EM created:" + emCreated.length + " modified-open:" + emModifiedOpen.length);
return {
success: true,
syncedAt: stamp,
range: { created: createdRange, modified: modifiedRange },
fetched: {
whatsapp: waRows.length,
callTickets: ctRows.length,
careEmails: emRows.length,
created: { whatsapp: waCreated.length, callTickets: ctCreated.length, careEmails: emCreated.length },
modifiedOpen: { whatsapp: waModifiedOpen.length, callTickets: ctModifiedOpen.length, careEmails: emModifiedOpen.length }
}
};
} catch (e) {
Logger.log("syncDevRevYesterdayIST FAILED: " + e.message);
throw e;
}
}

function syncDevRevDayWindowIST() {
var createdRange = getISTRangeForTodayWindow(8, 0, 16, 31);
var modifiedRange = getISTRangeForTodayWindow(8, 0, 16, 31);
Logger.log("syncDevRevDayWindowIST created_date after=" + createdRange.after + " before=" + createdRange.before);
Logger.log("syncDevRevDayWindowIST modified_date after=" + modifiedRange.after + " before=" + modifiedRange.before + " (open-only)");

var ss = SpreadsheetApp.getActiveSpreadsheet();
try {
var waCreated = fetchDevRevWhatsApp(createdRange.after, createdRange.before);
var ctCreated = fetchDevRevCallTickets(createdRange.after, createdRange.before);
var emCreated = fetchDevRevCareEmails(createdRange.after, createdRange.before);

var waModifiedOpen = fetchDevRevWhatsApp(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });
var ctModifiedOpen = fetchDevRevCallTickets(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });
var emModifiedOpen = fetchDevRevCareEmails(modifiedRange.after, modifiedRange.before, { dateField: "modified", onlyOpen: true });

var waRows = waCreated.concat(waModifiedOpen);
var ctRows = ctCreated.concat(ctModifiedOpen);
var emRows = emCreated.concat(emModifiedOpen);

writeDevRevToSheet(ss, SHEET_NAMES.whatsapp, waRows, COLS.whatsapp);
writeDevRevToSheet(ss, SHEET_NAMES.callTkts, ctRows, COLS.callTkts);
writeDevRevToSheet(ss, SHEET_NAMES.careEmails, emRows, COLS.careEmails);

var props = PropertiesService.getScriptProperties();
var stamp = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
props.setProperty("devrev_last_day_window_sync", stamp);
props.setProperty("devrev_last_day_window_label", Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd"));

buildDashboardCache();

Logger.log("syncDevRevDayWindowIST complete — WA created:" + waCreated.length + " modified-open:" + waModifiedOpen.length + " | CT created:" + ctCreated.length + " modified-open:" + ctModifiedOpen.length + " | EM created:" + emCreated.length + " modified-open:" + emModifiedOpen.length);
return {
success: true,
syncedAt: stamp,
range: { created: createdRange, modified: modifiedRange },
fetched: {
whatsapp: waRows.length,
callTickets: ctRows.length,
careEmails: emRows.length,
created: { whatsapp: waCreated.length, callTickets: ctCreated.length, careEmails: emCreated.length },
modifiedOpen: { whatsapp: waModifiedOpen.length, callTickets: ctModifiedOpen.length, careEmails: emModifiedOpen.length }
}
};
} catch (e) {
Logger.log("syncDevRevDayWindowIST FAILED: " + e.message);
throw e;
}
}

// ================================================================
// DEVREV — ONE CALENDAR DAY (IST), e.g. all items created on 2026-04-10 IST
// Run from editor: syncDevRevForCalendarDayIST("2026-04-10")
// ================================================================
function syncDevRevForCalendarDayIST(dateStr, opts) {
opts = opts || {};
var skipCacheRebuild = opts.skipCacheRebuild === true;
var appendOnly = opts.appendOnly === true;
var range = getISTCreatedRangeForCalendarDay(dateStr);
Logger.log("syncDevRevForCalendarDayIST " + (range.label || dateStr) + " after=" + range.after + " before=" + range.before);

var ss = SpreadsheetApp.getActiveSpreadsheet();
try {
var waRows = fetchDevRevWhatsApp(range.after, range.before);
var ctRows = fetchDevRevCallTickets(range.after, range.before);
var emRows = fetchDevRevCareEmails(range.after, range.before);

if (appendOnly) {
appendDevRevRowsToSheet(ss, SHEET_NAMES.whatsapp, waRows, COLS.whatsapp);
appendDevRevRowsToSheet(ss, SHEET_NAMES.callTkts, ctRows, COLS.callTkts);
appendDevRevRowsToSheet(ss, SHEET_NAMES.careEmails, emRows, COLS.careEmails);
} else {
writeDevRevToSheet(ss, SHEET_NAMES.whatsapp, waRows, COLS.whatsapp);
writeDevRevToSheet(ss, SHEET_NAMES.callTkts, ctRows, COLS.callTkts);
writeDevRevToSheet(ss, SHEET_NAMES.careEmails, emRows, COLS.careEmails);
}

if (!skipCacheRebuild) buildDashboardCache();

Logger.log("syncDevRevForCalendarDayIST complete — WA:" + waRows.length + " CT:" + ctRows.length + " EM:" + emRows.length);
return {
success: true,
day: range.label || dateStr,
range: range,
fetched: { whatsapp: waRows.length, callTickets: ctRows.length, careEmails: emRows.length }
};
} catch (e) {
Logger.log("syncDevRevForCalendarDayIST FAILED: " + e.message);
throw e;
}
}

/** Menu helper: April 10, 2026 (IST). Change the date string in code for another day. */
function syncDevRevApril10_2026_IST() {
return syncDevRevForCalendarDayIST("2026-04-10");
}

function formatISTDateOnly(d) {
return Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
}

function nextISTDate(dateStr) {
var parts = String(dateStr).split("-");
var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
d.setDate(d.getDate() + 1);
return Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
}

function compareISTDateStrings(a, b) {
if (!a && !b) return 0;
if (!a) return -1;
if (!b) return 1;
var at = new Date(String(a).trim() + "T00:00:00+05:30").getTime();
var bt = new Date(String(b).trim() + "T00:00:00+05:30").getTime();
if (isNaN(at) && isNaN(bt)) return 0;
if (isNaN(at)) return -1;
if (isNaN(bt)) return 1;
return at - bt;
}

function extractISTDateFromCell(val) {
if (!val) return "";
if (val instanceof Date && !isNaN(val.getTime())) {
return Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd");
}
var s = String(val).trim();
if (!s) return "";
var m = s.match(/^(\d{4}-\d{2}-\d{2})/);
if (m) return m[1];
var parsed = new Date(s);
if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, "Asia/Kolkata", "yyyy-MM-dd");
return "";
}

function getLatestBackfillCreatedDate(sheetName) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetName);
if (!sheet) return "";
var data = sheet.getDataRange().getValues();
if (data.length < 2) return "";
var headers = data[0].map(function(h) { return String(h).trim(); });
var createdIdx = headers.indexOf("Created date");
if (createdIdx === -1) return "";
for (var i = data.length - 1; i >= 1; i--) {
var d = extractISTDateFromCell(data[i][createdIdx]);
if (d) return d;
}
return "";
}

function getBackfillResumeDate(startDateStr) {
var props = PropertiesService.getScriptProperties();
var start = String(startDateStr || "2026-01-01").trim();
var recorded = String(props.getProperty(BACKFILL_PROGRESS_KEY) || "").trim();
var latestSheetDate = "";
[SHEET_NAMES.whatsapp, SHEET_NAMES.callTkts, SHEET_NAMES.careEmails].forEach(function(sheetName) {
var d = getLatestBackfillCreatedDate(sheetName);
if (compareISTDateStrings(d, latestSheetDate) > 0) latestSheetDate = d;
});
var base = start;
if (compareISTDateStrings(recorded, base) > 0) base = recorded;
if (compareISTDateStrings(latestSheetDate, base) > 0) base = latestSheetDate;
return nextISTDate(base);
}

function devrevBackfillSheetCols(sheetName) {
if (sheetName === SHEET_NAMES.whatsapp) return COLS.whatsapp;
if (sheetName === SHEET_NAMES.callTkts) return COLS.callTkts;
if (sheetName === SHEET_NAMES.careEmails) return COLS.careEmails;
throw new Error("Unsupported DevRev backfill sheet: " + sheetName);
}

function fullBackfillDevRevChronological(startDateStr, endDateStr, opts) {
opts = opts || {};
var skipCacheRebuild = opts.skipCacheRebuild === true;
var pauseMs = typeof opts.pauseMs === "number" ? opts.pauseMs : 100;
var startStr = String(startDateStr || "2026-01-01").trim();
var endStr = String(endDateStr || formatISTDateOnly(new Date())).trim();
var props = PropertiesService.getScriptProperties();
var totals = { whatsapp: 0, callTickets: 0, careEmails: 0, days: 0 };

Logger.log("Starting chronological backfill from " + startStr + " to " + endStr);
props.deleteProperty("devrev_last_sync");

var cur = startStr;
while (cur <= endStr) {
var result = syncDevRevForCalendarDayIST(cur, { skipCacheRebuild: true, appendOnly: true });
totals.whatsapp += result.fetched.whatsapp || 0;
totals.callTickets += result.fetched.callTickets || 0;
totals.careEmails += result.fetched.careEmails || 0;
totals.days++;
Logger.log("Chronological backfill day complete: " + cur + " — WA:" + result.fetched.whatsapp + " CT:" + result.fetched.callTickets + " EM:" + result.fetched.careEmails);
props.setProperty(BACKFILL_PROGRESS_KEY, cur);
cur = nextISTDate(cur);
if (pauseMs > 0) Utilities.sleep(pauseMs);
}

var now = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
props.setProperty("devrev_last_sync", now);
props.setProperty(BACKFILL_PROGRESS_KEY, endStr);
if (!skipCacheRebuild) buildDashboardCache();

return {
success: true,
mode: "chronological_backfill",
from: startStr,
to: endStr,
syncedAt: now,
days: totals.days,
fetched: {
whatsapp: totals.whatsapp,
callTickets: totals.callTickets,
careEmails: totals.careEmails
}
};
}

function resetDevRevBackfillSheets() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
[SHEET_NAMES.whatsapp, SHEET_NAMES.callTkts, SHEET_NAMES.careEmails].forEach(function(sheetName) {
var sheet = ss.getSheetByName(sheetName);
if (!sheet) throw new Error("Missing sheet: " + sheetName);
var cols = devrevBackfillSheetCols(sheetName);
sheet.clearContents();
sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
});
Logger.log("Reset DevRev backfill sheets to headers only");
}

// ================================================================
// DEVREV — FULL BACKFILL (run once to load all historical data)
// Resumes from the last completed backfill day and fetches forward to today
// ================================================================
function fullBackfillDevRev() {
var startDate = getBackfillResumeDate("2026-01-01");
var today = formatISTDateOnly(new Date());
if (compareISTDateStrings(startDate, today) > 0) {
try { SpreadsheetApp.getUi().alert("Full Backfill", "Backfill is already complete through " + today + ".", SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
return { success: true, skipped: true, startDate: startDate, endDate: today };
}
var result = fullBackfillDevRevChronological(startDate, today, { skipCacheRebuild: true, pauseMs: 100 });
var msg = "Fetched chronologically by Created date from " + startDate + " to now:\n" +
" Resumed from: " + startDate + "\n" +
" WhatsApp chats: " + result.fetched.whatsapp + "\n" +
" Call Tickets: " + result.fetched.callTickets + "\n" +
" Care Emails: " + result.fetched.careEmails + "\n" +
" Days processed: " + result.days + "\n\n" +
"Cache rebuild skipped for speed. Run Force Rebuild Cache after backfill completes.";
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
var cResp = devrevPostFullUrl(DEVREV_INTERNAL_CONVERSATIONS_LIST, { limit: 1, subtype: ["dealer_support"] });
var conv = (cResp.conversations || [])[0];
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

function truncateCacheText(val, maxLen) {
if (val === null || val === undefined) return "";
var s = String(val);
if (s.length <= maxLen) return s;
return s.substring(0, maxLen) + "…";
}

function compactRowsForCache(rows, sheetKey) {
var out = [];
var key = String(sheetKey || "");
for (var i = 0; i < rows.length; i++) {
var src = rows[i];
var row = {};
Object.keys(src).forEach(function(col) {
if (col === "Record JSON") return;
var v = src[col];
if (key === "careEmails" && col === "Body") v = truncateCacheText(v, 1200);
if ((key === "callTkts" || key === "whatsapp" || key === "careEmails") && (col === "Comments" || col === "Items")) v = truncateCacheText(v, 600);
row[col] = v;
});
out.push(row);
}
return out;
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
calls: compactRowsForCache(readSlimSheet(SHEET_NAMES.calls, CACHE_COLS.calls), "calls"),
callTkts: compactRowsForCache(readSlimSheet(SHEET_NAMES.callTkts, CACHE_COLS.callTkts), "callTkts"),
whatsapp: compactRowsForCache(readSlimSheet(SHEET_NAMES.whatsapp, CACHE_COLS.whatsapp), "whatsapp"),
careEmails: compactRowsForCache(readSlimSheet(SHEET_NAMES.careEmails, CACHE_COLS.careEmails), "careEmails"),
breaks: compactRowsForCache(readSlimSheet(SHEET_NAMES.breaks, CACHE_COLS.breaks), "breaks"),
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
if (["buildDashboardCache","onChangeHandler","autoRebuildAfterPaste","syncDevRevNow","syncDevRevHourly","syncDevRevYesterdayIST","syncDevRevDayWindowIST"].indexOf(fn) !== -1) {
ScriptApp.deleteTrigger(t); deleted++;
}
});
if (deleted > 0) Logger.log("Removed " + deleted + " old trigger(s)");

// 1. DevRev: import items created yesterday + open items modified yesterday — daily at 8:00 IST
ScriptApp.newTrigger("syncDevRevYesterdayIST").timeBased().everyDays(1).atHour(8).nearMinute(0).create();
Logger.log("Daily DevRev yesterday-import trigger created (8:00)");

// 2. DevRev: import items created today between 08:00 and 16:30 + open items modified in same window — daily at 16:30 IST
ScriptApp.newTrigger("syncDevRevDayWindowIST").timeBased().everyDays(1).atHour(16).nearMinute(30).create();
Logger.log("Daily DevRev day-window trigger created (16:30)");

// 3. onChange — auto-detect Ozonetel CSV paste
ScriptApp.newTrigger("onChangeHandler")
.forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
.onChange()
.create();

Logger.log("All triggers created successfully. Check View > Logs for confirmation.");

// Try to show UI alert — works when run from sheet menu, silently skips when run from editor
try {
SpreadsheetApp.getUi().alert(
"All Triggers Created ✓",
"Active schedule (project time zone — set to Asia/Kolkata for IST):\n\n" +
" 8:00 AM daily — DevRev: yesterday created + yesterday modified open items\n" +
" 4:30 PM daily — DevRev: today 8:00 to 4:30 window + modified open items\n" +
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
.addItem("📅 Sync DevRev — yesterday only (IST)", "syncDevRevYesterdayIST")
.addItem("📆 Sync DevRev — Apr 10, 2026 (IST)", "syncDevRevApril10_2026_IST")
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
"Ozonetel DevRev Tickets": ["Created date","Close date","Modified date","Last Agent Message Timestamp"],
"WhatsApp Chats": ["Created date","Modified date"],
"Care Emails": ["Created date","Close date","Modified date","Last Agent Message Timestamp"],
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
