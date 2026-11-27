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
  return main + email;
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

const root = path.join(__dirname, '..');
const main = fs.readdirSync(path.join(root, 'src')).filter(f => f.endsWith('.gs'))
  .map(f => fs.readFileSync(path.join(root, 'src', f), 'utf8')).join('\n');
const email = fs.readFileSync(path.join(root, 'email-monitor', 'EmailMonitor.gs'), 'utf8');
const count = runRegressionChecks(main, email, source =>
  vm.runInNewContext('(function () {\n' + source + '\n})()'));
console.log(count + ' status and email policy assertions passed; no external requests.');
