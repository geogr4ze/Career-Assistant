const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

/**
 * Exercise status refresh and email policies against isolated, mocked services.
 * @param {string} mainSource Concatenated Apps Script source.
 * @param {string} emailSource Standalone email source.
 * @param {Function} run Execute source in an isolated JavaScript environment.
 * @returns {number} Number of passing assertions.
 * @throws {Error} If any expected result differs.
 * @effects No external requests or production writes.
 */
function runRegressionChecks(mainSource, emailSource, run) {
  const main = run(mainSource + "\n" + mainScenarios());
  const email = run(emailSource + "\n" + emailScenarios());
  return main + email + run(mainSource + "\n" + pipelineFixtures() + mainPipeline())
    + run(emailSource + "\n" + pipelineFixtures() + emailPipeline());
}

/**
 * Build isolated Twilio and ledger regression scenarios.
 * @returns {string} JavaScript test body.
 * @effects None.
 */
function mainScenarios() {
  return `
let checks = 0;
function check(value, label) {
  if (!value) throw Error(label);
  checks++;
}
const now = new Date();
const sheetRows = [];
const calls = [];
const logs = [];
const outcomes = new Map();
const Logger = { log: message => logs.push(message) };
const SpreadsheetApp = { flush() {} };
credentials_ = () => ({});
log_ = (...args) => logs.push(args.join(" "));
twilioRequest_ = (c, path, method) => {
  check(method === "get", "status refresh must not send");
  const sid = path.split("/")[1].split(".")[0];
  calls.push(sid);
  const result = outcomes.get(sid);
  if (result === "throw") throw Error("mock transport");
  return result || {state:"ACCEPTED", sid, status:"delivered"};
};
const sheet = {
  getDataRange: () => ({ getValues: () => [[], ...sheetRows] }),
  getRange: (row, column) => ({
    setValue: value => { sheetRows[row - 2][column - 1] = value; },
    setValues: values => values[0].forEach((value, i) => {
      sheetRows[row - 2][column - 1 + i] = value;
    })
  })
};
book_ = () => ({getSheetByName: () => sheet});
function add(id, channel, status, age = 1000, note = "", sidOverride = "") {
  const sid = sidOverride || (channel === "SMS" ? "SM" : "CA") + String(id).padStart(32,"0");
  const time = new Date(now.getTime() - age);
  sheetRows.push(["E" + id,channel,"TEST","","ACCEPTED",time,time,sid,status,note,time]);
  return sid;
}
function refresh() {
  refreshStatuses_(sheet, notificationIndex_());
}
function reset() {
  sheetRows.length = 0;
  calls.length = 0;
  outcomes.clear();
}
check(!terminalStatus_("SMS","sent"),"SMS sent is not terminal");
check(terminalStatus_("SMS","undelivered"),"SMS undelivered is terminal");
check(terminalStatus_("CALL","completed"),"CALL completed is terminal");
let sid = add(1,"SMS","sent");
outcomes.set(sid,{state:"ACCEPTED",sid,status:"undelivered"});
refresh();
check(sheetRows[0][8] === "undelivered","sent can progress to undelivered");
refresh();
check(calls.length === 1,"terminal SMS not looked up again");
reset();
sid = add(2,"SMS","queued");
outcomes.set(sid,"throw");
refresh(); refresh(); refresh();
check(sheetRows[0][4] === "REVIEW","three failures require review");
check(sheetRows[0][8] === "queued","preserve last known status");
refresh();
check(calls.length === 3,"no further lookup after review");
reset();
sid = add(3,"CALL","queued");
outcomes.set(sid,{state:"FAILED",status:"HTTP 404"});
refresh();
check(sheetRows[0][4] === "REVIEW","missing provider resource requires review");
reset();
sid = add(4,"SMS","queued");
outcomes.set(sid,"throw"); refresh();
outcomes.delete(sid); refresh();
check(!sheetRows[0][9].includes("status failures"),"successful GET resets failure count");
reset();
add(5,"SMS","sent",73*60*60*1000);
refresh();
check(calls.length === 0 && sheetRows[0][4] === "REVIEW","stale status bounded");
reset();
for (let i=10;i<26;i++) {
  sid = add(i,"SMS","queued",60000);
  outcomes.set(sid,"throw");
}
const newer = add(99,"SMS","sent",1000);
refresh();
check(calls.length === 15,"15-request cap");
refresh();
check(calls.includes(newer),"old failures cannot starve newer rows");
reset();
add(101,"SMS","queued"); add(101,"SMS","queued"); add(102,"SMS","queued");
const index = notificationIndex_();
check(index.get("E101|SMS").duplicate,"duplicate main key quarantined");
localDay_ = () => "today";
check(countAttempts_(index,"SMS",now) === 3,"all duplicate attempts counted");
refresh();
check(calls.length === 1,"unrelated status refresh continues");
return checks;
`;
}

/**
 * Build isolated email duplicate and quiet-hour regression scenarios.
 * @returns {string} JavaScript test body.
 * @effects None.
 */
function emailScenarios() {
  return `
let checks = 0;
function check(value,label) {
  if (!value) throw Error(label);
  checks++;
}
const console = { warn() {}, error() {} };
const quiet = {start:"21:00",end:"08:00"};
check(emInQuiet_("21:00",quiet),"quiet start inclusive");
check(emInQuiet_("02:00",quiet),"overnight quiet");
check(!emInQuiet_("08:00",quiet),"quiet end exclusive");
check(!emInQuiet_("20:59",quiet),"daytime permitted");
check(emInQuiet_("12:00",{start:"11:00",end:"13:00"}),"same-day window");
const rows = [
  ["Event ID","State","Attempted day (PT)","Updated (PT)","Note"],
  ["A","SENT","PT:2026-11-26"],["A","REVIEW","PT:2026-11-26"],["B","BASELINE",""]
];
const idx = emIndex_({getDataRange:()=>({getDisplayValues:()=>rows})});
check(idx.get("A").duplicate,"duplicate email key quarantined");
check(idx.has("B"),"unrelated event remains indexed");
check(idx.attemptDays.length===2,"all duplicate email attempts counted");
const settings = [["Setting","Value"],["QUIET_START","21:00"],["QUIET_END","08:00"]];
const book = {getSheetByName:()=>({getDataRange:()=>({
  getDisplayValues:()=>settings,getValues:()=>settings
})})};
check(emQuietSettings_(book).start==="21:00","read shared quiet hours");
settings.push(["QUIET_START","22:00"]);
let rejected=false;
try {emQuietSettings_(book);} catch (_) {rejected=true;}
check(rejected,"duplicate quiet settings fail closed");
check(emPriority_("NVIDIA Corp."),"priority suffix preserved");
check(!emPriority_("Not NVIDIA"),"no false employer match");
return checks;
`;
}

/**
 * Construct synthetic tracker and provider mocks.
 * @returns {string} Isolated test source.
 * @effects None; constructs fixtures only.
 */
function pipelineFixtures() {
  return `
let checks = 0;
/** Verify a test condition and count successful assertions; throws on failure. */
function check(value, label) {
  if (!value) throw Error(label);
  checks++;
}
/** Create a mutable in-memory sheet; does not access Google Sheets. */
function fakeSheet(rows) {
  return {
    rows,
    getDataRange: () => ({
      getValues: () => rows.map(row => row.slice()),
      getDisplayValues: () => rows.map(row => row.map(value =>
        value instanceof Date ? value.toISOString() : String(value ?? "")))
    }),
    getRange: (row, col, height = 1, width = 1) => ({
      getDisplayValues: () => rows.slice(row - 1, row - 1 + height).map(values =>
        values.slice(col - 1, col - 1 + width).map(String)),
      setValue: value => { rows[row - 1][col - 1] = value; },
      setValues: values => values.forEach((line, i) => line.forEach((value, j) => {
        rows[row - 1 + i][col - 1 + j] = value;
      }))
    }),
    appendRow: row => rows.push(row.slice()),
    getLastRow: () => rows.length
  };
}
const applicationHeaders = Array(24).fill("");
applicationHeaders[0] = "Application ID";
applicationHeaders[1] = "Employer";
applicationHeaders[2] = "Position";
applicationHeaders[10] = "Status";
applicationHeaders[11] = "Date submitted";
applicationHeaders[13] = "Confirmation / proof";
const application = Array(24).fill("");
application[0] = "SIM-APP";
application[1] = "NVIDIA";
application[2] = "Intern";
application[10] = "Submitted";
application[11] = new Date();
application[13] = "Synthetic confirmation";
const responseHeaders = ["Response ID", "Application ID", "Received date/time (PT)",
  "Sender", "Response type", "Subject", "Summary", "Gmail thread / source URL",
  "Action required", "Action deadline (PT)", "Action status"];
const response = ["SIM-REPLY", "SIM-APP", new Date(), "", "Interview invitation", "",
  "Synthetic interview", "", "Choose time", "", ""];
const sheets = {
  Applications: fakeSheet([applicationHeaders, application]),
  Responses: fakeSheet([responseHeaders, response]),
  Settings: fakeSheet([["Setting", "Value"], ["QUIET_START", "21:00"], ["QUIET_END", "08:00"]]),
  "Approval Queue": fakeSheet([["Approval ID", "Application ID", "", "", "", "", "Status"]]),
  "System Logs": fakeSheet([["Timestamp (PT)", "Level", "Operation", "Message"]])
};
const book = { getSheetByName: name => sheets[name] };
const properties = {
  CAREER_LIVE_ARMED: "YES",
  EMAIL_MONITOR_ENABLED: "YES",
  EMAIL_MONITOR_FROM: "alias@example.invalid",
  EMAIL_MONITOR_TO: "owner@example.invalid",
  EMAIL_MONITOR_SHEET_ID: "SYNTHETIC"
};
let clock = "12:00";
const PropertiesService = {
  getScriptProperties: () => ({ getProperty: key => properties[key] })
};
const LockService = {
  getScriptLock: () => ({ tryLock: () => true, releaseLock() {} })
};
const SpreadsheetApp = { openById: () => book, flush() {} };
const Utilities = { formatDate: (date, zone, format) =>
  format === "HH:mm" ? clock : "2026-11-26" };
const Logger = { log() {} };
const console = { warn() {}, error() {} };
`;
}

/**
 * Exercise automatic SMS and call delivery against mocks.
 * @returns {string} Isolated test source.
 * @effects None; constructs fixtures only.
 */
function mainPipeline() {
  return `
sheets.Notifications = fakeSheet([CA.sheets.Notifications]);
book_ = () => book;
const config = {
  MODE: "LIVE", QUIET_START: "21:00", QUIET_END: "08:00",
  MAX_SMS_PER_DAY: "20", MAX_CALLS_PER_DAY: "3",
  ...Object.fromEntries(CA.routes.map(key => ["ROUTE_" + key, "OFF"])),
  ROUTE_SUBMITTED: "SMS", ROUTE_INTERVIEW: "BOTH"
};
getSettings_ = () => config;
credentials_ = () => ({});
const sends = [];
deliver_ = (channel, type) => {
  const active = sheets.Notifications.rows.filter(r => r[4] === "SENDING");
  check(active.length === 1 && active[0][6] instanceof Date
    && active[0][1] === channel && active[0][2] === type,
    "persist the correct attempt before provider call");
  sends.push([channel, type]);
  return { state:"ACCEPTED", sid:(channel === "SMS" ? "SM" : "CA")
    + String(sends.length).padStart(32,"0"), status:"queued", note:"mock accepted" };
};
twilioRequest_ = (c, path, method) => {
  check(method === "get", "status lookup is read-only");
  const sid = path.split("/")[1].split(".")[0];
  return { state:"ACCEPTED", sid, status:sid.startsWith("CA") ? "completed" : "delivered" };
};
clock = "02:00";
pollCareerAlerts();
check(sends.length === 0, "quiet hours block SMS and calls");
check(sheets.Notifications.rows.length === 4, "three events queued by channel");
clock = "08:00";
pollCareerAlerts();
check(sends.length === 3, "one submission SMS, one interview SMS and call");
check(sends.some(s => s[0] === "CALL" && s[1] === "INTERVIEW"), "interview call routed");
pollCareerAlerts();
check(sends.length === 3, "repeated polls do not resend");
check(sheets.Notifications.rows.slice(1).every(r => r[4] === "ACCEPTED"),
  "accepted requests recorded");
return checks;
`;
}

/**
 * Exercise automatic email delivery against mocks.
 * @returns {string} Isolated test source.
 * @effects None; constructs fixtures only.
 */
function emailPipeline() {
  return `
sheets["Email Notifications"] = fakeSheet([EM.headers]);
const Session = { getEffectiveUser: () => ({getEmail: () => properties.EMAIL_MONITOR_TO}) };
const MailApp = { getRemainingDailyQuota: () => 100 };
const sends = [];
const GmailApp = {
  getAliases: () => [properties.EMAIL_MONITOR_FROM],
  sendEmail: (to, subject, text, options) => {
    const active = sheets["Email Notifications"].rows.filter(r => r[1] === "SENDING");
    check(active.length === 1, "persist email attempt before Gmail call");
    check(to === properties.EMAIL_MONITOR_TO && options.from === properties.EMAIL_MONITOR_FROM,
      "email targets verified owner from verified alias");
    sends.push({subject,text,options});
  }
};
clock = "02:00";
pollEmployerMonitorEmail();
check(sends.length === 0, "quiet hours block email");
clock = "08:00";
pollEmployerMonitorEmail();
check(sends.length === 2, "submission and interview email routed");
check(sends.every(s => s.subject.startsWith("PRIORITY")), "priority HTML email subjects");
pollEmployerMonitorEmail();
check(sends.length === 2, "repeat email poll does not resend");
check(sheets["Email Notifications"].rows.slice(1).every(r => r[1] === "SENT"),
  "Gmail acceptance recorded");
return checks;
`;
}

const root = path.join(__dirname, '..');
const main = fs.readdirSync(path.join(root, 'src')).filter(f => f.endsWith('.gs'))
  .map(f => fs.readFileSync(path.join(root, 'src', f), 'utf8')).join('\n');
const email = fs.readFileSync(path.join(root, 'email-monitor', 'EmailMonitor.gs'), 'utf8');
const count = runRegressionChecks(main, email, source =>
  vm.runInNewContext('(function () {\n' + source + '\n})()'));
console.log(count + ' status and email policy assertions passed; no external requests.');
