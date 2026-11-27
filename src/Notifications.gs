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
 * Index ledger rows without hiding duplicate delivery attempts.
 * @returns {Map} Event/channel keys with representative rows and duplicate flags.
 * @effects Reads Notifications; records all physical rows for daily attempt accounting.
 * @throws {Error} If the notification ledger is unavailable.
 */
function notificationIndex_() {
  const s = book_().getSheetByName('Notifications');
  const rows = s.getDataRange().getValues().slice(1);
  const map = new Map();
  map.attemptRows = rows;
  rows.forEach((r, i) => {
    if (!r[0] || !r[1])
      return;
    const key = r[0] + '|' + r[1];
    const prior = map.get(key);
    if (prior) {
      prior.duplicate = true;
      return;
    }
    map.set(key, { row: i + 2, values: r, duplicate: false });
  });
  if ([...map.values()].some(item => item.duplicate))
    Logger.log('Duplicate notification keys quarantined; inspect Notifications.');
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
 * @returns {{state:string,sid:string,status:string,note:string}} Delivery request outcome;
 *   exceptions from deliver_ become REVIEW.
 * @effects Writes and flushes ledger state before/after a potentially billable Twilio request.
 *   Does not retry ambiguous outcomes.
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
  const rows = idx.attemptRows || Array.from(idx.values(), item => item.values);
  rows.forEach(r => {
    if (r[1] === ch && r[6] instanceof Date && localDay_(r[6]) === localDay_(now))
      n++;
  });
  return n;
}

/**
 * Preview up to 25 routed events that have no ledger row or are still PENDING.
 * @requires Initialized tracker and spreadsheet UI context.
 * @effects Reads settings/events/ledger, writes a diagnostic log and opens a dialog. Does not call
 *   Twilio or enforce current quiet-hour/usage eligibility.
 * @returns {void}
 */
function previewCareerAlerts() {
  const c = getSettings_();
  const idx = notificationIndex_();
  const planned = [];
  eventList_().forEach(e => channels_(c['ROUTE_' + e.type]).forEach(ch => {
    const old = idx.get(e.id + '|' + ch);
    if (!old || (!old.duplicate && old.values[4] === 'PENDING'))
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
 * @effects Acquires the script lock, writes ledger rows, may make paid requests, and refreshes
 *   provider statuses. Quiet hours, daily limits and a ten-attempt per-run cap restrict sends.
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
        if (item && (item.duplicate || item.values[4] !== 'PENDING'))
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
 * Check whether the provider status is terminal for this specific channel.
 * @param {string} channel - SMS or CALL.
 * @param {string} status - Last observed provider status.
 * @returns {boolean} True only for a known terminal status.
 * @effects None.
 */
function terminalStatus_(channel, status) {
  const value = String(status).toLowerCase();
  if (channel === 'SMS')
    return ['delivered', 'undelivered', 'failed', 'canceled', 'received'].includes(value);
  if (channel === 'CALL')
    return ['completed', 'busy', 'no-answer', 'failed', 'canceled'].includes(value);
  return false;
}

/**
 * Recover the persisted count of consecutive failed GET requests.
 * @param {string} note - Notification ledger note.
 * @returns {number} Failure count, or zero when no marker is present.
 * @effects None.
 */
function lookupFailures_(note) {
  const match = String(note || '').match(/\[status failures: (\d+)\]/);
  return match ? Number(match[1]) : 0;
}

/**
 * Remove the internal status retry marker from an existing diagnostic note.
 * @param {string} note - Notification ledger note.
 * @returns {string} Original diagnostic text without the retry marker.
 * @effects None.
 */
function lookupBaseNote_(note) {
  return String(note || '').replace(/\s*\[status failures: \d+\].*$/, '').trim();
}

/**
 * Refresh unresolved delivery records, prioritizing the least recently checked.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} s - Notifications sheet.
 * @param {Map} idx - Ledger index from notificationIndex_.
 * @requires Caller holds the script lock and has valid Twilio credentials.
 * @effects Makes at most 15 GET requests; persists every lookup outcome.
 * Stops tracking after three consecutive errors, a missing resource, or 72 hours.
 * A REVIEW state preserves the last known provider status and never causes a resend.
 * @returns {void} Does not send SMS or make calls.
 * @throws {Error} If credentials or ledger writes fail.
 */
function refreshStatuses_(s, idx) {
  const c = credentials_();
  const now = new Date();
  const maxAge = 72 * 60 * 60 * 1000;
  const candidates = [...idx.values()].filter(({ values: r, duplicate }) =>
    !duplicate && r[4] === 'ACCEPTED'
    && ['SMS', 'CALL'].includes(r[1]) && !terminalStatus_(r[1], r[8]));
  candidates.sort((a, b) => {
    const left = a.values[10] instanceof Date ? a.values[10].getTime() : 0;
    const right = b.values[10] instanceof Date ? b.values[10].getTime() : 0;
    return (left || 0) - (right || 0) || a.row - b.row;
  });

  for (const { row, values: r } of candidates.slice(0, 15)) {
    const channel = r[1];
    const sid = String(r[7] || '');
    const validSid = channel === 'SMS' ? /^(SM|MM)[0-9a-f]{32}$/i : /^CA[0-9a-f]{32}$/i;
    const attempted = r[6] instanceof Date ? r[6].getTime() : NaN;
    const baseNote = lookupBaseNote_(r[9]);
    let review = '';
    let status = r[8];
    let note = baseNote;

    if (!validSid.test(sid)) {
      review = 'Missing or invalid provider SID.';
    } else if (!Number.isFinite(attempted) || now.getTime() - attempted >= maxAge) {
      review = 'Status unresolved after 72 hours or attempt timestamp unavailable.';
    } else {
      let result;
      try {
        const prefix = channel === 'SMS' ? 'Messages/' : 'Calls/';
        result = twilioRequest_(c, prefix + sid + '.json', 'get');
      } catch (_) {
        // No raw exception text: it can contain provider credentials or URLs.
        result = { status: 'Transport error' };
      }

      if (result.state === 'ACCEPTED' && result.sid === sid
        && result.status && result.status !== 'accepted') {
        status = result.status;
        // Successful lookups reset the consecutive failure counter.
      } else {
        const failures = lookupFailures_(r[9]) + 1;
        const detail = /^HTTP \d{3}$/.test(result.status) ? result.status : 'Lookup unavailable';
        note = baseNote + ' [status failures: ' + failures + '] ' + detail;
        if (result.status === 'HTTP 404' || result.status === 'HTTP 410')
          review = 'Provider resource unavailable (' + result.status + ').';
        else if (failures >= 3)
          review = 'Three consecutive status lookups failed (' + detail + ').';
      }
    }

    if (review) {
      // Persist the review state before diagnostics; never infer actual delivery failure.
      s.getRange(row, 5).setValue('REVIEW');
      note = baseNote + ' ' + review + ' Inspect provider logs; no automatic retry or resend.';
    }
    s.getRange(row, 9, 1, 3).setValues([[status, note.trim(), now]]);
    SpreadsheetApp.flush();
    if (review)
      log_('ERROR', 'Status lookup', channel + ' row ' + row + ': ' + review);
  }
}

/**
 * Ask for explicit permission to send one test notification, including while the tracker is in DRY_RUN.
 * @param {string} ch - SMS or CALL.
 * @requires Valid settings, configured credentials and a spreadsheet UI context.
 * @effects On approval, locks execution, checks quiet hours/daily limits, creates a TEST row and
 *   attempts a potentially billable request.
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
