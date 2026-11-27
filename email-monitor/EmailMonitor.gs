const EM = Object.freeze({
  "ledger": "Email Notifications",
  "headers": [
    "Event ID",
    "State",
    "Attempted day (PT)",
    "Updated (PT)",
    "Note"
  ],
  "zone": "America/Vancouver",
  "maxPerDay": 30,
  "maxPerRun": 10,
  "responseTypes": [
    "Acknowledgment only",
    "Recruiter reply",
    "Assessment",
    "Interview invitation",
    "Offer",
    "Rejection",
    "Other substantive reply",
    "Progression"
  ]
});

/**
 * Escape untrusted tracker data for HTML.
 * @param {*} value Input.
 * @returns {string} Escaped text.
 */
function emEscape_(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  })[c]);
}

/**
 * Normalize and bound a tracker field.
 * @param {*} value Input.
 * @param {number} limit Maximum length.
 * @returns {string} Single-line text.
 */
function emText_(value, limit = 240) {
  return String(value == null ? "" : value).replace(/\s+/g, " ").trim().slice(0, limit);
}

/**
 * Recognize explicitly named major technology employers.
 * @param {string} employer Employer name.
 * @returns {boolean} Whether the employer matches the priority list.
 */
function emPriority_(employer) {
  return /^(apple|microsoft|google|alphabet|amazon|aws|amazon web services|meta|nvidia|amd|intel|qualcomm|samsung|ibm|oracle|cisco|salesforce|adobe|tesla|spacex|openai)( inc\.?| corporation| corp\.?| canada)?$/i
    .test(employer);
}

/**
 * Extract deduplicated events from tracker rows; no I/O.
 * @param {Array[]} apps Application rows.
 * @param {Array[]} responses Response rows.
 * @returns {Object[]} Verified events with priority employers first.
 */
function emEvents_(apps, responses) {
  const byId = new Map(apps.filter(r => r[0]).map(r => [String(r[0]), r]));
  const events = new Map();
  for (const a of apps) {
    if (!a[0] || !a[1] || !a[2] || !a[11] || !a[13]) continue;
    const id = "APP:" + a[0] + ":SUBMITTED";
    events.set(id, {id, employer:emText_(a[1]), role:emText_(a[2]),
      status:"Application submitted", received:emText_(a[11]), deadline:"",
      action:"No action unless requested.", summary:"Submission confirmation recorded."});
  }
  for (const r of responses) {
    const a = byId.get(String(r[1]));
    if (!r[0] || !a || !a[1] || !a[2] || !r[2] || !EM.responseTypes.includes(r[4])) continue;
    const id = "RESPONSE:" + r[0];
    events.set(id, {id, employer:emText_(a[1]), role:emText_(a[2]),
      status:emText_(r[4]), received:emText_(r[2]), deadline:emText_(r[9]),
      action:emText_(r[8], 500), summary:emText_(r[6], 500)});
  }
  return Array.from(events.values()).sort((a,b) => Number(emPriority_(b.employer))
    - Number(emPriority_(a.employer)));
}

/**
 * Render a notification with escaped tracker data.
 * @param {Object} event Verified tracker event.
 * @param {string} dashboard Trusted tracker URL.
 * @returns {Object} Subject, plain text and HTML.
 */
function emMessage_(event, dashboard) {
  const priority = emPriority_(event.employer);
  const subject = (priority ? "PRIORITY | " : "Career Assistant | ")
    + event.employer + " — " + event.status;
  const fields = [["Employer",event.employer],["Position",event.role],["Status",event.status],
    ["Received (PT)",event.received],["Deadline",event.deadline],["Action",event.action]]
    .filter(([,value]) => value);
  const rows = fields.map(([label,value]) =>
    '<tr><td style="padding:9px 12px;color:#64748b;vertical-align:top">'
    + emEscape_(label) + '</td><td style="padding:9px 12px;font-weight:600">'
    + emEscape_(value) + '</td></tr>').join("");
  const html = '<div style="max-width:580px;margin:auto;font-family:Arial,sans-serif;'
    + 'color:#243247;border:1px solid #dce3ed;border-radius:12px;overflow:hidden">'
    + '<div style="background:#142338;color:white;padding:24px">'
    + '<strong style="font-size:20px;letter-spacing:1px">CAREER ASSISTANT</strong>'
    + '<div style="font-size:13px;margin-top:6px;color:#c5d1e1">Employer Response Monitor</div></div>'
    + '<div style="padding:24px">'
    + (priority ? '<div style="color:#995b13;font-weight:bold">PRIORITY — TECH GIANT</div>' : '')
    + '<h2 style="margin:12px 0">' + emEscape_(event.status) + '</h2>'
    + '<p>Sir, here is your latest application update.</p>'
    + (event.summary ? '<p style="line-height:1.5">' + emEscape_(event.summary) + '</p>' : '')
    + '<table style="width:100%;border-collapse:collapse">' + rows + '</table>'
    + '<p style="margin-top:24px"><a style="background:#142338;color:white;text-decoration:none;'
    + 'display:inline-block;padding:12px 18px;border-radius:6px" href="'
    + emEscape_(dashboard) + '">View application dashboard</a></p></div>'
    + '<div style="padding:16px 24px;border-top:1px solid #dce3ed;color:#64748b;font-size:12px">'
    + 'Career Assistant &middot; Automated notification</div></div>';
  return {subject, html, text:["Career Assistant — Employer Response Monitor",event.summary,
    ...fields.map(([k,v]) => k+": "+v), dashboard].filter(Boolean).join("\n")};
}

/**
 * Load private settings and verify alias and recipient.
 * @returns {Object} Validated configuration.
 * @throws {Error} If configuration is invalid.
 */
function emConfig_() {
  const props = PropertiesService.getScriptProperties();
  const sender = props.getProperty("EMAIL_MONITOR_FROM");
  const to = props.getProperty("EMAIL_MONITOR_TO");
  const id = props.getProperty("EMAIL_MONITOR_SHEET_ID");
  const owner = Session.getEffectiveUser().getEmail();
  if (!sender || !to || !id || !owner || to.toLowerCase() !== owner.toLowerCase())
    throw Error("Set private sender, tracker ID and recipient (executing Gmail account).");
  if (!GmailApp.getAliases().some(alias => alias.toLowerCase() === sender.toLowerCase()))
    throw Error("Configured sender is not a verified Gmail alias.");
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw Error("Invalid tracker ID.");
  return {props,sender,to,id,dashboard:"https://docs.google.com/spreadsheets/d/"+id+"/edit"};
}

/**
 * Read a tab and verify headers.
 * @param {Object} book Tracker spreadsheet.
 * @param {string} name Tab name.
 * @param {Object} headers Expected column headers by index.
 * @param {number[]} dateColumns Columns requiring native valid date values.
 * @returns {Array[]} Display rows.
 * @throws {Error} If schema differs.
 */
function emRows_(book, name, headers, dateColumns = []) {
  const sheet=book.getSheetByName(name);
  if (!sheet) throw Error("Missing tab: "+name);
  const range = sheet.getDataRange();
  const rows = range.getDisplayValues();
  const raw = range.getValues();
  for (const [i,header] of Object.entries(headers))
    if (rows[0][Number(i)] !== header) throw Error("Unexpected header in "+name+": "+header);
  return rows.slice(1).map((row, index) => {
    for (const col of dateColumns) {
      const value = raw[index + 1][col];
      if (!(value instanceof Date) || isNaN(value.getTime())) row[col] = "";
    }
    return row;
  });
}

/**
 * Read and extract application events.
 * @param {Object} book Tracker spreadsheet.
 * @returns {Object[]} Verified events.
 * @throws {Error} If tracker reads fail.
 */
function emReadEvents_(book) {
  return emEvents_(emRows_(book,"Applications",{0:"Application ID",1:"Employer",
    2:"Position",11:"Date submitted",13:"Confirmation / proof"}, [11]),
  emRows_(book,"Responses",{0:"Response ID",1:"Application ID",2:"Received date/time (PT)",
    4:"Response type",6:"Summary",8:"Action required",9:"Action deadline (PT)"}, [2]));
}

/**
 * Open a dedicated ledger; create only during explicit setup.
 * @param {Object} book Tracker spreadsheet.
 * @param {boolean} create Allow creation.
 * @returns {Object} Ledger sheet.
 * @throws {Error} If schema differs.
 */
function emLedger_(book,create) {
  let sheet=book.getSheetByName(EM.ledger);
  if (!sheet && create) {
    sheet=book.insertSheet(EM.ledger);
    sheet.appendRow(EM.headers);
    sheet.setFrozenRows(1);
  }
  if (!sheet || sheet.getRange(1,1,1,EM.headers.length).getDisplayValues()[0].join("|")
    !== EM.headers.join("|")) throw Error("Missing or invalid email notification ledger.");
  return sheet;
}

/**
 * Index ledger entries without changing them.
 * @param {Object} sheet Ledger sheet.
 * @returns {Map} Event IDs and states.
 * @throws {Error} On duplicate IDs.
 */
function emIndex_(sheet) {
  const index=new Map();
  sheet.getDataRange().getDisplayValues().slice(1).forEach((r,i)=>{
    if (!r[0]) return;
    if (index.has(r[0])) throw Error("Duplicate email ledger event ID.");
    index.set(r[0],{row:i+2,state:r[1],day:r[2]});
  });
  return index;
}

/**
 * Format a Vancouver calendar day.
 * @param {Date} date Instant.
 * @returns {string} Text-only date key (PT:yyyy-MM-dd).
 */
function emDay_(date) {
  // Prefix the key so Sheets cannot auto-convert it into a locale-specific date.
  return "PT:" + Utilities.formatDate(date, EM.zone, "yyyy-MM-dd");
}

/**
 * Baseline old events and explicitly enable future email delivery.
 * @requires Private properties, verified alias and tracker access.
 * @effects Creates ledger and installs trigger; sends no email.
 * @throws {Error} If initialization fails.
 */
function enableEmployerMonitorEmail() {
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const c=emConfig_();
    if (c.props.getProperty("EMAIL_MONITOR_ENABLED")==="YES")
      throw Error("Email monitor is already enabled.");
    const book=SpreadsheetApp.openById(c.id);
    const events=emReadEvents_(book);
    const ledger=emLedger_(book,true);
    const idx=emIndex_(ledger);
    for (const e of events)
      if (!idx.has(e.id)) ledger.appendRow([e.id,"BASELINE","",new Date(),"Existing event skipped"]);
    SpreadsheetApp.flush();
    ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==="pollEmployerMonitorEmail")
      .forEach(t=>ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger("pollEmployerMonitorEmail").timeBased().everyMinutes(5).create();
    c.props.setProperty("EMAIL_MONITOR_ENABLED","YES");
  } finally {
    lock.releaseLock();
  }
}

/**
 * Disable the email trigger and preserve the ledger.
 * @effects Changes only this standalone project; does not touch SMS or calls.
 */
function disableEmployerMonitorEmail() {
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    PropertiesService.getScriptProperties().deleteProperty("EMAIL_MONITOR_ENABLED");
    ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==="pollEmployerMonitorEmail")
      .forEach(t=>ScriptApp.deleteTrigger(t));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Process new verified tracker events once automatically.
 * @requires Enabled project, valid alias and ledger.
 * @effects Records attempts before sending; caps usage; never retries uncertain sends.
 * @throws {Error} If configuration or ledger access fails.
 */
function pollEmployerMonitorEmail() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    if (PropertiesService.getScriptProperties().getProperty("EMAIL_MONITOR_ENABLED")!=="YES") return;
    const c=emConfig_(),book=SpreadsheetApp.openById(c.id);
    const ledger=emLedger_(book,false),idx=emIndex_(ledger);
    const day=emDay_(new Date());
    let remaining=Math.max(0,EM.maxPerDay-Array.from(idx.values()).filter(x=>x.day===day).length);
    let attempts=0;
    for (const e of emReadEvents_(book)) {
      if (idx.has(e.id)) continue;
      if (remaining<=0 || attempts>=EM.maxPerRun || MailApp.getRemainingDailyQuota()<1) break;
      if (c.props.getProperty("EMAIL_MONITOR_ENABLED")!=="YES") return;
      const message=emMessage_(e,c.dashboard);
      ledger.appendRow([e.id,"SENDING",day,new Date(),"Attempt recorded before Gmail call"]);
      const row=ledger.getLastRow();
      SpreadsheetApp.flush();
      remaining--;
      attempts++;
      try {
        GmailApp.sendEmail(c.to,message.subject,message.text,
          {from:c.sender,name:"Career Assistant",htmlBody:message.html});
      } catch (_) {
        ledger.getRange(row,2,1,4).setValues([["REVIEW",day,new Date(),
          "Send outcome uncertain; inspect Sent folder. No automatic retry."]]);
        SpreadsheetApp.flush();
        console.error("Email send outcome uncertain; inspect Email Notifications.");
        continue;
      }
      ledger.getRange(row,2,1,4).setValues([["SENT",day,new Date(),
        "Gmail accepted send; inbox delivery not independently verified."]]);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}
