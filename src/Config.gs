const CA = Object.freeze({
  zone: 'America/Vancouver',
  sheets: {
    Settings: ['Setting', 'Value', 'Description'],
    Notifications: [
      'Event ID',
      'Channel',
      'Event type',
      'Application ID',
      'State',
      'Created (PT)',
      'Attempted (PT)',
      'Provider SID',
      'Provider status',
      'Note',
      'Updated (PT)'
    ],
    'Approval Queue': [
      'Approval ID',
      'Application ID',
      'Created (PT)',
      'Question',
      'Options',
      'Deadline (PT)',
      'Status',
      'User decision',
      'Source'
    ],
    'System Logs': ['Timestamp (PT)', 'Level', 'Operation', 'Message']
  },
  routes: [
    'SUBMITTED',
    'ACKNOWLEDGMENT',
    'EMPLOYER_REPLY',
    'PROGRESSION',
    'ASSESSMENT',
    'INTERVIEW',
    'OFFER',
    'REJECTION',
    'APPROVAL'
  ],
  labels: {
    SUBMITTED: 'An application was submitted.',
    ACKNOWLEDGMENT: 'An employer acknowledged your application.',
    EMPLOYER_REPLY: 'An employer replied.',
    PROGRESSION: 'An application advanced to the next stage.',
    ASSESSMENT: 'An assessment invitation arrived.',
    INTERVIEW: 'An interview invitation arrived.',
    OFFER: 'A job offer arrived.',
    REJECTION: 'An application was rejected.',
    APPROVAL: 'Your approval is required.',
    TEST: 'This is a test of your career assistant.'
  }
});

/**
 * Build default rows for the Settings tab. New installations start unarmed, with all routes OFF
 *   and limits/quiet hours awaiting user input.
 * @returns {Array<Array<string>>} Setting name, initial value and description for each row.
 * @effects None; does not write to Sheets.
 */
function settingsDefaults_() {
  return [
    ['MODE', 'DRY_RUN', 'DRY_RUN or LIVE. Use the menu to arm live alerts.'],
    ['QUIET_START', '', 'Start of quiet hours in 24-hour HH:mm (e.g. 21:00).'],
    ['QUIET_END', '', 'End of quiet hours in 24-hour HH:mm (e.g. 08:00).'],
    [
      'MAX_SMS_PER_DAY', '', 'Positive integer; attempts per Vancouver calendar day.'
    ],
    [
      'MAX_CALLS_PER_DAY', '', 'Positive integer; attempts per Vancouver calendar day.'
    ],
    [
      'POLL_MINUTES', '5', 'Choose 1, 5, 10, 15 or 30; reinstall trigger after changing.'
    ],
    ...CA.routes.map(k => ['ROUTE_' + k, 'OFF', 'Choose OFF, SMS, CALL or BOTH.'])
  ];
}

/**
 * Open the tracker using the private spreadsheet ID saved during setup.
 * @requires Script Properties contains CAREER_SHEET_ID and the caller can access that spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The tracker.
 * @throws {Error} If setup has not saved an ID or the spreadsheet cannot be opened.
 */
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('CAREER_SHEET_ID');
  if (!id)
    throw Error('Run setupCareerAssistant from the bound spreadsheet first.');
  return SpreadsheetApp.openById(id);
}

/**
 * Read the Settings tab into a lookup object; displayed values are trimmed strings.
 * @requires The tracker has a Settings tab with names in column A and values in B.
 * @returns {Object<string,string>} Settings keyed by name; ignores blank names.
 * @throws {Error} If a setting name is duplicated or the sheet is unavailable.
 */
function getSettings_() {
  const rows = book_().getSheetByName('Settings').getDataRange().getDisplayValues().slice(1);
  const config = {};
  rows.forEach(([k, v]) => {
    if (k) {
      if (Object.prototype.hasOwnProperty.call(config, k))
        throw Error('Duplicate setting: ' + k);
      config[k] = String(v).trim();
    }
  });
  return config;
}

/**
 * Replace the value of an existing setting without inserting a new row.
 * @param {string} key - Exact setting name.
 * @param {string} value - New value to write.
 * @effects Writes column B of the first matching setting row.
 * @throws {Error} If the named setting does not exist or Sheets access fails.
 */
function setSetting_(key, value) {
  const s = book_().getSheetByName('Settings');
  const rows = s.getDataRange().getDisplayValues();
  const i = rows.findIndex(r => r[0] === key);
  if (i < 1)
    throw Error('Missing setting: ' + key);
  s.getRange(i + 1, 2).setValue(value);
}

/**
 * Validate the settings required for alert routing and delivery limits. Polling interval is validated separately.
 * @param {Object<string,string>} c - Settings returned by getSettings_.
 * @returns {void} Returns normally only if mode, quiet hours, limits and all routes pass validation.
 * @throws {Error} For an invalid mode, equal/invalid quiet times, limits outside 1–100, or an unsupported route.
 * @effects None.
 */
function configValid_(c) {
  if (!['LIVE', 'DRY_RUN'].includes(c.MODE))
    throw Error('MODE must be LIVE or DRY_RUN.');
  if (!validTime_(c.QUIET_START) || !validTime_(c.QUIET_END) || c.QUIET_START === c.QUIET_END)
    throw Error('Choose distinct QUIET_START and QUIET_END in HH:mm format.');
  for (const k of ['MAX_SMS_PER_DAY', 'MAX_CALLS_PER_DAY'])
    if (!/^[1-9]\d*$/.test(c[k] || '') || Number(c[k]) > 100)
      throw Error(k + ' must be a positive integer up to 100.');
  for (const k of CA.routes)
    if (!['OFF', 'SMS', 'CALL', 'BOTH'].includes(c['ROUTE_' + k]))
      throw Error('Invalid ROUTE_' + k);
}

/**
 * Check the 24-hour HH:mm format using a regular expression.
 * @param {string} t - Candidate time string.
 * @returns {boolean} True for hours 00–23 and minutes 00–59; otherwise false.
 */
function validTime_(t) {
  return /^(?:[01]\d|2[0-3]):[0-5]\d$/.test(t || '');
}

/**
 * Determine whether the current clock time is inside the quiet window, including windows crossing midnight.
 * @param {string} now - Current local clock in HH:mm.
 * @param {string} start - Quiet-window start in HH:mm (inclusive).
 * @param {string} end - Quiet-window end in HH:mm (exclusive).
 * @requires All times are zero-padded valid HH:mm strings; start and end differ.
 * @returns {boolean} Whether alerts are currently restricted.
 */
function inQuiet_(now, start, end) {
  return start < end ? now >= start && now < end : now >= start || now < end;
}

/**
 * Expand a route into the individual notification channels.
 * @param {string} route - OFF, SMS, CALL or BOTH.
 * @returns {string[]} SMS then CALL for BOTH, a single channel for SMS/CALL, or an empty array otherwise.
 */
function channels_(route) {
  return route === 'BOTH' ? ['SMS', 'CALL'] : ['SMS', 'CALL'].includes(route) ? [route] : [];
}

/**
 * Format an instant as a Vancouver calendar date for daily usage accounting.
 * @param {Date} date - Valid date/time.
 * @returns {string} Local date in yyyy-MM-dd format.
 */
function localDay_(date) {
  return Utilities.formatDate(date, CA.zone, 'yyyy-MM-dd');
}

/**
 * Format an instant as a Vancouver clock time for quiet-hour checks.
 * @param {Date} date - Valid date/time.
 * @returns {string} Local time in HH:mm format.
 */
function localClock_(date) {
  return Utilities.formatDate(date, CA.zone, 'HH:mm');
}
