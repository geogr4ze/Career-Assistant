const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),path=require('path');
const root=path.join(__dirname,'../src');const ctx=vm.createContext({});
for(const f of fs.readdirSync(root).filter(x=>x.endsWith('.gs')))new vm.Script(fs.readFileSync(path.join(root,f),'utf8'),{filename:f}).runInContext(ctx);
/**
 * Evaluate a JavaScript expression in the isolated VM containing the Apps Script functions.
 * @param {string} s - Test expression to evaluate.
 * @returns {*} Value produced by the expression.
 * @throws {Error} If parsing or evaluation fails; assertions may deliberately test errors.
 */
const e=s=>vm.runInContext(s,ctx);assert.equal(e('validTime_("08:00")'),true);assert.equal(e('validTime_("25:00")'),false);assert.equal(e('inQuiet_("23:00","21:00","08:00")'),true);assert.equal(e('inQuiet_("07:59","21:00","08:00")'),true);assert.equal(e('inQuiet_("08:00","21:00","08:00")'),false);assert.equal(e('inQuiet_("21:00","21:00","08:00")'),true);assert.equal(e('inQuiet_("12:00","11:00","13:00")'),true);assert.equal(e('channels_("BOTH").join(",")'),'SMS,CALL');assert.equal(e('channels_("OFF").length'),0);assert.equal(e('responseType_("Interview invitation")'),'INTERVIEW');assert.equal(e('responseType_("unknown")'),null);assert.equal(e('xmlEscape_("a&<b>")'),'a&amp;&lt;b&gt;');assert.match(e('notificationText_("OFFER")'),/Boss/);assert.throws(()=>e('configValid_({MODE:"DRY_RUN"})'));assert.equal(e('settingsDefaults_().filter(r=>r[0].startsWith("ROUTE_")).every(r=>r[1]==="OFF")'),true);
console.log('15 local checks passed; all Apps Script files parsed. No external API calls made.');
const trial = '{mode:"TRIAL",from:"SENDER_PLACEHOLDER",to:"RECIPIENT_PLACEHOLDER"}';
const prod = '{mode:"PRODUCTION",from:"SENDER_PLACEHOLDER",to:"RECIPIENT_PLACEHOLDER"}';
assert.equal(e(`twilioPayload_(${trial},"SMS","TEST").payload.Body`), 'sms_internal_alerts');
assert.match(e(`twilioPayload_(${trial},"CALL","TEST").payload.Url`), /voice_text_to_speech$/);
assert.equal(e(`twilioPayload_(${trial},"CALL","TEST").payload.Twiml`), undefined);
assert.match(e(`twilioPayload_(${prod},"SMS","TEST").payload.Body`), /Boss/);
assert.match(e(`twilioPayload_(${prod},"CALL","TEST").payload.Twiml`), /<Say>Boss/);
assert.throws(()=>e('twilioPayload_({mode:"INVALID"},"SMS","TEST")'));
console.log('6 additional trial/production payload checks passed.');
