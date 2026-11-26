const CA = Object.freeze({
  zone: 'America/Vancouver',
  sheets: {
    Settings: ['Setting','Value','Description'],
    Notifications: ['Event ID','Channel','Event type','Application ID','State','Created (PT)','Attempted (PT)','Provider SID','Provider status','Note','Updated (PT)'],
    'Approval Queue': ['Approval ID','Application ID','Created (PT)','Question','Options','Deadline (PT)','Status','User decision','Source'],
    'System Logs': ['Timestamp (PT)','Level','Operation','Message']
  },
  routes: ['SUBMITTED','ACKNOWLEDGMENT','EMPLOYER_REPLY','PROGRESSION','ASSESSMENT','INTERVIEW','OFFER','REJECTION','APPROVAL'],
  labels: {SUBMITTED:'An application was submitted.',ACKNOWLEDGMENT:'An employer acknowledged your application.',EMPLOYER_REPLY:'An employer replied.',PROGRESSION:'An application advanced to the next stage.',ASSESSMENT:'An assessment invitation arrived.',INTERVIEW:'An interview invitation arrived.',OFFER:'A job offer arrived.',REJECTION:'An application was rejected.',APPROVAL:'Your approval is required.',TEST:'This is a test of your career assistant.'}
});
function settingsDefaults_() {
  return [
    ['MODE','DRY_RUN','DRY_RUN or LIVE. Use the menu to arm live alerts.'],
    ['QUIET_START','','Start of quiet hours in 24-hour HH:mm (e.g. 21:00).'],
    ['QUIET_END','','End of quiet hours in 24-hour HH:mm (e.g. 08:00).'],
    ['MAX_SMS_PER_DAY','','Positive integer; attempts per Vancouver calendar day.'],
    ['MAX_CALLS_PER_DAY','','Positive integer; attempts per Vancouver calendar day.'],
    ['POLL_MINUTES','5','Choose 1, 5, 10, 15 or 30; reinstall trigger after changing.'],
    ...CA.routes.map(k=>['ROUTE_'+k,'OFF','Choose OFF, SMS, CALL or BOTH.'])
  ];
}
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('CAREER_SHEET_ID');
  if (!id) throw Error('Run setupCareerAssistant from the bound spreadsheet first.');
  return SpreadsheetApp.openById(id);
}
function getSettings_() {
  const rows = book_().getSheetByName('Settings').getDataRange().getDisplayValues().slice(1);
  const config = {};
  rows.forEach(([k,v])=>{ if (k) {if (Object.prototype.hasOwnProperty.call(config,k)) throw Error('Duplicate setting: '+k);config[k]=String(v).trim();}});
  return config;
}
function setSetting_(key,value) {
  const s=book_().getSheetByName('Settings');
  const rows=s.getDataRange().getDisplayValues();
  const i=rows.findIndex(r=>r[0]===key);
  if(i<1) throw Error('Missing setting: '+key);
  s.getRange(i+1,2).setValue(value);
}
function configValid_(c) {
  if (!['LIVE','DRY_RUN'].includes(c.MODE)) throw Error('MODE must be LIVE or DRY_RUN.');
  if (!validTime_(c.QUIET_START)||!validTime_(c.QUIET_END)||c.QUIET_START===c.QUIET_END) throw Error('Choose distinct QUIET_START and QUIET_END in HH:mm format.');
  for (const k of ['MAX_SMS_PER_DAY','MAX_CALLS_PER_DAY']) if(!/^[1-9]\d*$/.test(c[k]||'')||Number(c[k])>100) throw Error(k+' must be a positive integer up to 100.');
  for(const k of CA.routes) if(!['OFF','SMS','CALL','BOTH'].includes(c['ROUTE_'+k])) throw Error('Invalid ROUTE_'+k);
}
function validTime_(t){return /^(?:[01]\d|2[0-3]):[0-5]\d$/.test(t||'');}
function inQuiet_(now,start,end){return start<end ? now>=start&&now<end : now>=start||now<end;}
function channels_(route){return route==='BOTH'?['SMS','CALL']:['SMS','CALL'].includes(route)?[route]:[];}
function localDay_(date){return Utilities.formatDate(date,CA.zone,'yyyy-MM-dd');}
function localClock_(date){return Utilities.formatDate(date,CA.zone,'HH:mm');}
