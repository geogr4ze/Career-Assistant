/**
 * Append an operational log entry to the System Logs tab.
 * @param {string} level - Severity label, such as INFO or ERROR.
 * @param {string} op - Operation name.
 * @param {string} message - Safe diagnostic text; callers should not include secrets.
 * @effects Adds one timestamped row to System Logs.
 */
function log_(level, op, message) {
  book_().getSheetByName('System Logs').appendRow([new Date(), level, op, message]);
}

/**
 * Index existing notification rows by event ID and channel for duplicate checks.
 * @returns {Map<string,{row:number,values:Array<*>}>} Composite key to sheet row number and raw values; later duplicate rows replace earlier ones.
 * @effects Reads Notifications without modifying it.
 */
function notificationIndex_() {
  const s = book_().getSheetByName('Notifications');
  const rows = s.getDataRange().getValues();
  const map = new Map();
  rows.slice(1).forEach((r, i) => {
    if (r[0] && r[1])
      map.set(r[0] + '|' + r[1], { row: i + 2, values: r });
  });
  return map;
}

/**
 * Create a notification ledger row before attempting delivery.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} s - Notifications sheet.
 * @param {{id:string,type:string,app:string}} e - Event record.
 * @param {string} ch - SMS or CALL.
 * @param {string} state - Initial ledger state.
 * @param {Date} now - Creation/update timestamp.
 * @param {string} note - Optional diagnostic description; falsy values become empty.
 * @returns {number} Row number after appending.
 * @effects Appends a row; provider and attempt fields initially remain blank.
 */
function appendNotification_(s, e, ch, state, now, note) {
  s.appendRow([e.id, ch, e.type, e.app, state, now, '', '', '', note || '', now]);
  return s.getLastRow();
}

/**
 * Attempt one notification after recording SENDING and its attempt timestamp.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} s - Notifications sheet.
 * @param {number} row - Existing ledger row number (1-based).
 * @param {{id:string,type:string,app:string}} e - Event record.
 * @param {string} ch - SMS or CALL.
 * @requires Caller has checked delivery policy and holds the script lock.
 * @returns {{state:string,sid:string,status:string,note:string}} Delivery request outcome; exceptions from deliver_ become REVIEW.
 * @effects Writes and flushes ledger state before/after a potentially billable Twilio request. Does not retry ambiguous outcomes.
 */
function attempt_(s, row, e, ch) {
  // Persist the attempt before calling the provider so an interrupted run is not blindly retried.
  const now = new Date();
  s.getRange(row, 5).setValue('SENDING');
  s.getRange(row, 7).setValue(now);
  SpreadsheetApp.flush();
  let result;
  try {
    result = deliver_(ch, e.type);
  } catch (_) {
    result = {
      state: 'REVIEW',
      sid: '',
      status: 'Unknown',
      note: 'Request outcome uncertain; inspect Twilio logs. No automatic retry.'
    };
  }
  s.getRange(row, 5).setValue(result.state);
  s.getRange(row, 8, 1, 4).setValues([[result.sid, result.status, result.note, new Date()]]);
  SpreadsheetApp.flush();
  return result;
}

/**
 * Count recorded attempts for a channel on the current Vancouver calendar day.
 * @param {Map<string,{row:number,values:Array<*>}>} idx - Notification index.
 * @param {string} ch - SMS or CALL.
 * @param {Date} now - Instant identifying the local day.
 * @returns {number} Count of matching rows with a valid Date attempt field; includes unsuccessful attempts.
 */
function countAttempts_(idx, ch, now) {
  let n = 0;
  idx.forEach(({ values: r }) => {
    if (r[1] === ch && r[6] instanceof Date && localDay_(r[6]) === localDay_(now))
      n++;
  });
  return n;
}

/**
 * Preview up to 25 routed events that have no ledger row or are still PENDING.
 * @requires Initialized tracker and spreadsheet UI context.
 * @effects Reads settings/events/ledger, writes a diagnostic log and opens a dialog. Does not call Twilio or enforce current quiet-hour/usage eligibility.
 * @returns {void}
 */
function previewCareerAlerts() {
  const c = getSettings_();
  const idx = notificationIndex_();
  const planned = [];
  eventList_().forEach(e => channels_(c['ROUTE_' + e.type]).forEach(ch => {
    const old = idx.get(e.id + '|' + ch);
    if (!old || old.values[4] === 'PENDING')
      planned.push(ch + ' | ' + e.type + ' | ' + e.app);
  }));
  const msg = planned.length
    ? planned.slice(0, 25).join('\n')
    : 'No eligible notifications under the current routes.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Dry-run preview', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Process eligible notification events on a timer; skip execution unless LIVE mode and explicit arming are both active.
 * @requires Initialized tracker, valid configuration and configured Twilio credentials for live operation.
 * @effects Acquires the script lock, writes ledger rows, may make paid requests, and refreshes provider statuses. Quiet hours, daily limits and a ten-attempt per-run cap restrict sends.
 * @returns {void} Returns early if the lock is unavailable or live delivery is disabled.
 * @throws {Error} Polling failures are logged and rethrown; the lock is released in finally.
 */
function pollCareerAlerts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000))
    return;
  try {
    // Both controls must be enabled; changing only the spreadsheet cell is insufficient.
    const c = getSettings_();
    if (c.MODE !== 'LIVE'
      || PropertiesService.getScriptProperties().getProperty('CAREER_LIVE_ARMED') !== 'YES')
      return;
    configValid_(c);
    credentials_();
    const now = new Date();
    const s = book_().getSheetByName('Notifications');
    const idx = notificationIndex_();
    // Include failed attempts when calculating the remaining daily budget.
    const counts = { SMS: countAttempts_(idx, 'SMS', now), CALL: countAttempts_(idx, 'CALL', now) };
    let attempted = 0;
    for (const e of eventList_())
      for (const ch of channels_(c['ROUTE_' + e.type])) {
        // An event/channel pair already in a non-PENDING state is not sent again.
        let item = idx.get(e.id + '|' + ch);
        if (item && item.values[4] !== 'PENDING')
          continue;
        if (!item) {
          const row = appendNotification_(s, e, ch, 'PENDING', now, 'Queued');
          item = { row, values: [e.id, ch, e.type, e.app, 'PENDING'] };
          idx.set(e.id + '|' + ch, item);
        }
        // Leave restricted notifications pending for a future polling run.
        if (inQuiet_(localClock_(now), c.QUIET_START, c.QUIET_END)
          || counts[ch] >= Number(c[ch === 'SMS' ? 'MAX_SMS_PER_DAY' : 'MAX_CALLS_PER_DAY'])
          || attempted >= 10)
          continue;
        // Recheck kill switch before each paid request.
        if (PropertiesService.getScriptProperties().getProperty('CAREER_LIVE_ARMED') !== 'YES'
          || getSettings_().MODE !== 'LIVE')
          return;
        attempt_(s, item.row, e, ch);
        counts[ch]++;
        attempted++;
      }
    // Refresh previously accepted requests separately from creating notifications.
    refreshStatuses_(s, idx);
  } catch (err) {
    log_('ERROR', 'Poll', String(err.message || 'Unknown error').slice(0, 200));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Refresh provider status for previously accepted notifications with nonterminal statuses.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} s - Notifications sheet.
 * @param {Map<string,{row:number,values:Array<*>}>} idx - Snapshot of existing ledger rows.
 * @effects Makes at most 15 eligible Twilio GET requests and updates provider status and timestamp on success. Individual request failures are ignored.
 * @returns {void} Does not initiate new SMS or calls.
 */
function refreshStatuses_(s, idx) {
  const c = credentials_();
  let n = 0;
  for (const { row, values: r } of idx.values()) {
    if (n >= 15)
      break;
    if (r[4] !== 'ACCEPTED' || !r[7] || [
      'delivered', 'undelivered', 'failed', 'completed', 'busy', 'no-answer', 'canceled', 'sent'
    ].includes(String(r[8])))
      continue;
    const prefix = r[1] === 'SMS' ? 'Messages/' : 'Calls/';
    if (!/^(SM|MM|CA)[0-9a-fA-F]{32}$/.test(r[7]))
      continue;
    try {
      const result = twilioRequest_(c, prefix + r[7] + '.json', 'get');
      if (result.sid) {
        s.getRange(row, 9).setValue(result.status);
        s.getRange(row, 11).setValue(new Date());
      }
    } catch (_) {
    }
    n++;
  }
}

/**
 * Ask for explicit permission to send one test notification, including while the tracker is in DRY_RUN.
 * @param {string} ch - SMS or CALL.
 * @requires Valid settings, configured credentials and a spreadsheet UI context.
 * @effects On approval, locks execution, checks quiet hours/daily limits, creates a TEST row and attempts a potentially billable request.
 * @throws {Error} If settings/credentials are invalid, the lock cannot be obtained or delivery is restricted.
 * @returns {void} Cancellation sends nothing.
 */
function manualTest_(ch) {
  const c = getSettings_();
  configValid_(c);
  credentials_();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Send paid test?',
    'Send one ' + ch + ' to the number stored privately in ALERT_TO_NUMBER? Charges or trial units may apply.',
    ui.ButtonSet.YES_NO
  ) !== ui.Button.YES)
    return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    const idx = notificationIndex_();
    if (inQuiet_(localClock_(now), c.QUIET_START, c.QUIET_END))
      throw Error('Test blocked during quiet hours.');
    if (countAttempts_(idx, ch, now) >= Number(c[ch === 'SMS' ? 'MAX_SMS_PER_DAY' : 'MAX_CALLS_PER_DAY']))
      throw Error('Daily limit reached.');
    const s = book_().getSheetByName('Notifications');
    const e = { id: 'TEST:' + Utilities.getUuid(), type: 'TEST', app: '' };
    const row = appendNotification_(s, e, ch, 'PENDING', now, 'Manually approved test.');
    const result = attempt_(s, row, e, ch);
    Logger.log(result.state + ' ' + result.status);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menu entry point for one manually approved SMS test.
 * @effects Delegates to manualTest_('SMS'); may incur a charge if approved and allowed.
 * @returns {void}
 */
function testCareerSms() {
  manualTest_('SMS');
}

/**
 * Menu entry point for one manually approved voice-call test.
 * @effects Delegates to manualTest_('CALL'); may incur a charge if approved and allowed.
 * @returns {void}
 */
function testCareerCall() {
  manualTest_('CALL');
}
