/**
 * Read a tracker tab, checking its expected column headers before interpreting its data.
 * @param {string} name - Sheet tab name.
 * @param {Array<Array<number|string>>} expected - Pairs of zero-based column index and exact header text.
 * @returns {Array<Array<*>>} Raw cell values for all rows after the header.
 * @throws {Error} If the tab is missing or a required header is different.
 */
function rows_(name, expected) {
  const s = book_().getSheetByName(name);
  if (!s)
    throw Error('Missing ' + name);
  const values = s.getDataRange().getValues();
  const headers = values[0].map(String);
  for (const [col, key] of expected)
    if (headers[col] !== key)
      throw Error('Unexpected ' + name + ' header ' + key);
  return values.slice(1);
}

/**
 * Translate a tracker response label into an internal alert event type.
 * @param {string} type - Exact response label from the Responses tab.
 * @returns {string|null} The corresponding event type, or null for an unrecognized label.
 */
function responseType_(type) {
  return {
    'Acknowledgment only': 'ACKNOWLEDGMENT',
    'Recruiter reply': 'EMPLOYER_REPLY',
    'Assessment': 'ASSESSMENT',
    'Interview invitation': 'INTERVIEW',
    'Offer': 'OFFER',
    'Rejection': 'REJECTION',
    'Other substantive reply': 'EMPLOYER_REPLY',
    'Progression': 'PROGRESSION'
  }[type] || null;
}

/**
 * Extract eligible submission, response and pending approval events from the tracker.
 * @requires Applications, Responses and Approval Queue have the expected headers and stable IDs.
 * @returns {Array<{id:string,type:string,app:string}>} Events with duplicate IDs removed; first occurrence wins.
 * @effects Reads Sheets only. Does not send alerts or update the ledger.
 * @throws {Error} If a required tab or header is missing.
 */
function eventList_() {
  const a = rows_('Applications', [
    [0, 'Application ID'], [10, 'Status'], [11, 'Date submitted'], [13, 'Confirmation / proof']
  ]);
  const r = rows_('Responses', [[0, 'Response ID'], [1, 'Application ID'], [4, 'Response type']]);
  const q = rows_('Approval Queue', [[0, 'Approval ID'], [1, 'Application ID'], [6, 'Status']]);
  // Build a lookup of known application IDs before accepting response records.
  const out = [];
  const apps = new Set(a.filter(x => x[0]).map(x => String(x[0])));
  // A submission needs an ID, a real date and recorded confirmation/proof.
  a.forEach(x => {
    if (x[0] && x[11] instanceof Date && !isNaN(x[11]) && x[13])
      out.push({ id: 'APP:' + x[0] + ':SUBMITTED', type: 'SUBMITTED', app: String(x[0]) });
  });
  // Accept only recognized responses linked to known applications.
  r.forEach(x => {
    const t = responseType_(String(x[4]));
    if (x[0] && x[1] && apps.has(String(x[1])) && t)
      out.push({ id: 'RESPONSE:' + x[0], type: t, app: String(x[1]) });
  });
  // Only unresolved approvals become approval events.
  q.forEach(x => {
    if (x[0] && x[6] === 'Pending')
      out.push({ id: 'APPROVAL:' + x[0], type: 'APPROVAL', app: String(x[1] || '') });
  });
  // Keep the first occurrence of each stable event ID.
  const seen = new Set();
  return out.filter(e => {
    if (seen.has(e.id))
      return false;
    seen.add(e.id);
    return true;
  });
}

/**
 * Build a generic alert message without including private employer or application details.
 * @param {string} type - Internal event type (including TEST).
 * @returns {string} Personalized greeting and generic event description; unknown types use a fallback.
 */
function notificationText_(type) {
  return 'Boss, ' + (CA.labels[type] || 'A job search update arrived.') + ' Please check your job search dashboard.';
}
