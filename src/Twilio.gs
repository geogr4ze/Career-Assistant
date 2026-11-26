/**
 * Read private Twilio settings from Script Properties and validate their basic format.
 * @returns {{sid:string,token:string,from:string,to:string,mode:string}} Account credentials, distinct E.164 numbers and explicit TRIAL/PRODUCTION mode.
 * @throws {Error} If properties are absent or invalid. Format validation does not verify account ownership or service availability.
 * @effects Reads private properties; does not log their values or contact Twilio.
 */
function credentials_() {
  const p = PropertiesService.getScriptProperties();
  const sid = p.getProperty('TWILIO_ACCOUNT_SID') || '';
  const token = p.getProperty('TWILIO_AUTH_TOKEN') || '';
  const from = p.getProperty('TWILIO_FROM_NUMBER') || '';
  const to = p.getProperty('ALERT_TO_NUMBER') || '';
  const mode = (p.getProperty('TWILIO_MODE') || '').trim().toUpperCase();
  if (!/^AC[0-9a-fA-F]{32}$/.test(sid)
    || !token
    || !/^\+[1-9]\d{7,14}$/.test(from)
    || !/^\+[1-9]\d{7,14}$/.test(to)
    || from === to) {
    throw Error('Configure all four private Script Properties with valid values.');
  }
  if (!['TRIAL', 'PRODUCTION'].includes(mode)) {
    throw Error('Set TWILIO_MODE explicitly to TRIAL or PRODUCTION.');
  }
  return { sid, token, from, to, mode };
}

/**
 * Escape XML special characters before embedding text in TwiML.
 * @param {*} s - Value converted to a string before escaping.
 * @returns {string} XML-escaped text.
 */
function xmlEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/**
 * Construct the correct Twilio resource path and request body for one alert.
 * @param {{mode:string,from:string,to:string}} c - Validated mode and phone numbers.
 * @param {string} ch - SMS or CALL.
 * @param {string} type - Internal event type used for production messages.
 * @returns {{path:string,payload:Object}} Trial template or production SMS/TwiML payload.
 * @throws {Error} If the mode or channel is unsupported.
 * @effects None; payload construction does not send a request.
 */
function twilioPayload_(c, ch, type) {
  if (!['TRIAL', 'PRODUCTION'].includes(c.mode))
    throw Error('Invalid Twilio mode.');
  // Trial accounts use provider-defined content; production uses the generated alert text.
  const base = { From: c.from, To: c.to };
  if (ch === 'SMS') {
    return {
      path: 'Messages.json', payload: {
        ...base,
        Body: c.mode === 'TRIAL' ? 'sms_internal_alerts' : notificationText_(type)
      }
    };
  }
  if (ch === 'CALL') {
    if (c.mode === 'TRIAL') {
      return {
        path: 'Calls.json', payload: {
          ...base,
          Url: 'https://webhooks.twilio.com/v1/Voice/Template/voice_text_to_speech'
        }
      };
    }
    return {
      path: 'Calls.json', payload: {
        ...base,
        Twiml: '<Response><Say>' + xmlEscape_(notificationText_(type)) + '</Say><Hangup/></Response>',
        Timeout: '20', TimeLimit: '45'
      }
    };
  }
  throw Error('Unknown notification channel');
}

/**
 * Make an authenticated request to the Twilio account API and normalize its response.
 * @param {{sid:string,token:string,mode:string}} c - Validated private credentials.
 * @param {string} path - Account-relative API resource path.
 * @param {string} method - HTTP method (get or post in current callers).
 * @param {Object} [payload] - Form data for requests that require it.
 * @returns {{state:string,sid:string,status:string,note:string}} ACCEPTED for a 2xx response with a SID, FAILED for 4xx, REVIEW otherwise.
 * @effects Contacts Twilio; POST requests can incur charges. Logs only sanitized provider error codes, not raw response bodies.
 * @throws {Error} Transport or Apps Script service exceptions may propagate; ACCEPTED does not guarantee delivery.
 */
function twilioRequest_(c, path, method, payload) {
  // Authenticate with the private account credentials and do not follow redirects.
  const opts = {
    method, headers: { Authorization: 'Basic ' + Utilities.base64Encode(c.sid + ':' + c.token) },
    muteHttpExceptions: true, followRedirects: false
  };
  if (payload)
    opts.payload = payload;
  const res = UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + c.sid + '/' + path, opts);
  const code = res.getResponseCode();
  let json = {};
  try {
    json = JSON.parse(res.getContentText()) || {};
  } catch (_) {
  }
  // A provider SID confirms acceptance of the request, not final delivery.
  if (code >= 200 && code < 300 && json.sid) {
    return {
      state: 'ACCEPTED', sid: json.sid, status: json.status || 'accepted',
      note: 'Provider accepted request (' + c.mode + '); delivery not guaranteed.'
    };
  }
  // Record only numeric provider codes, not response bodies or credentials.
  const errorCode = /^\d{3,9}$/.test(String(json.code || '')) ? String(json.code) : 'Unknown';
  const note = 'Twilio error code: ' + errorCode + ' | Mode: ' + c.mode + ' | Inspect Twilio logs before retrying.';
  Logger.log('Twilio HTTP ' + code + ' | ' + note);
  return {
    state: code >= 400 && code < 500 ? 'FAILED' : 'REVIEW', sid: '', status: 'HTTP ' + code, note
  };
}

/**
 * Send one SMS or call using the currently configured Twilio mode.
 * @param {string} ch - SMS or CALL.
 * @param {string} type - Internal notification event type.
 * @requires Caller has checked authorization, quiet hours and attempt limits.
 * @returns {{state:string,sid:string,status:string,note:string}} Normalized Twilio API result.
 * @effects Reads credentials and makes one potentially billable POST request.
 * @throws {Error} Invalid credentials, unsupported channel or transport failure.
 */
function deliver_(ch, type) {
  const c = credentials_();
  const request = twilioPayload_(c, ch, type);
  return twilioRequest_(c, request.path, 'post', request.payload);
}
