function credentials_() {
  const p = PropertiesService.getScriptProperties();
  const sid = p.getProperty('TWILIO_ACCOUNT_SID') || '';
  const token = p.getProperty('TWILIO_AUTH_TOKEN') || '';
  const from = p.getProperty('TWILIO_FROM_NUMBER') || '';
  const to = p.getProperty('ALERT_TO_NUMBER') || '';
  const mode = (p.getProperty('TWILIO_MODE') || '').trim().toUpperCase();
  if (!/^AC[0-9a-fA-F]{32}$/.test(sid) || !token || !/^\+[1-9]\d{7,14}$/.test(from) || !/^\+[1-9]\d{7,14}$/.test(to) || from === to) {
    throw Error('Configure all four private Script Properties with valid values.');
  }
  if (!['TRIAL', 'PRODUCTION'].includes(mode)) {
    throw Error('Set TWILIO_MODE explicitly to TRIAL or PRODUCTION.');
  }
  return {sid, token, from, to, mode};
}

function xmlEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function twilioPayload_(c, ch, type) {
  if (!['TRIAL', 'PRODUCTION'].includes(c.mode)) throw Error('Invalid Twilio mode.');
  const base = {From: c.from, To: c.to};
  if (ch === 'SMS') {
    return {path: 'Messages.json', payload: {...base,
      Body: c.mode === 'TRIAL' ? 'sms_internal_alerts' : notificationText_(type)}};
  }
  if (ch === 'CALL') {
    if (c.mode === 'TRIAL') {
      return {path: 'Calls.json', payload: {...base,
        Url: 'https://webhooks.twilio.com/v1/Voice/Template/voice_text_to_speech'}};
    }
    return {path: 'Calls.json', payload: {...base,
      Twiml: '<Response><Say>' + xmlEscape_(notificationText_(type)) + '</Say><Hangup/></Response>',
      Timeout: '20', TimeLimit: '45'}};
  }
  throw Error('Unknown notification channel');
}

function twilioRequest_(c, path, method, payload) {
  const opts = {method, headers: {Authorization: 'Basic ' + Utilities.base64Encode(c.sid + ':' + c.token)},
    muteHttpExceptions: true, followRedirects: false};
  if (payload) opts.payload = payload;
  const res = UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + c.sid + '/' + path, opts);
  const code = res.getResponseCode();
  let json = {};
  try { json = JSON.parse(res.getContentText()) || {}; } catch (_) {}
  if (code >= 200 && code < 300 && json.sid) {
    return {state: 'ACCEPTED', sid: json.sid, status: json.status || 'accepted',
      note: 'Provider accepted request (' + c.mode + '); delivery not guaranteed.'};
  }
  // Record only numeric provider codes, not response bodies or credentials.
  const errorCode = /^\d{3,9}$/.test(String(json.code || '')) ? String(json.code) : 'Unknown';
  const note = 'Twilio error code: ' + errorCode + ' | Mode: ' + c.mode + ' | Inspect Twilio logs before retrying.';
  Logger.log('Twilio HTTP ' + code + ' | ' + note);
  return {state: code >= 400 && code < 500 ? 'FAILED' : 'REVIEW', sid: '', status: 'HTTP ' + code, note};
}

function deliver_(ch, type) {
  const c = credentials_();
  const request = twilioPayload_(c, ch, type);
  return twilioRequest_(c, request.path, 'post', request.payload);
}
