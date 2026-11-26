function log_(level,op,message){book_().getSheetByName('System Logs').appendRow([new Date(),level,op,message]);}
function notificationIndex_(){const s=book_().getSheetByName('Notifications');const rows=s.getDataRange().getValues();const map=new Map();rows.slice(1).forEach((r,i)=>{if(r[0]&&r[1])map.set(r[0]+'|'+r[1],{row:i+2,values:r});});return map;}
function appendNotification_(s,e,ch,state,now,note){s.appendRow([e.id,ch,e.type,e.app,state,now,'','','',note||'',now]);return s.getLastRow();}
function attempt_(s,row,e,ch){
  const now=new Date();s.getRange(row,5).setValue('SENDING');s.getRange(row,7).setValue(now);SpreadsheetApp.flush();
  let result;try{result=deliver_(ch,e.type);}catch(_){result={state:'REVIEW',sid:'',status:'Unknown',note:'Request outcome uncertain; inspect Twilio logs. No automatic retry.'};}
  s.getRange(row,5).setValue(result.state);s.getRange(row,8,1,4).setValues([[result.sid,result.status,result.note,new Date()]]);SpreadsheetApp.flush();return result;
}
function countAttempts_(idx,ch,now){let n=0;idx.forEach(({values:r})=>{if(r[1]===ch&&r[6] instanceof Date&&localDay_(r[6])===localDay_(now))n++;});return n;}
function previewCareerAlerts(){const c=getSettings_();const idx=notificationIndex_();const planned=[];eventList_().forEach(e=>channels_(c['ROUTE_'+e.type]).forEach(ch=>{const old=idx.get(e.id+'|'+ch);if(!old||old.values[4]==='PENDING')planned.push(ch+' | '+e.type+' | '+e.app);}));
  const msg=planned.length?planned.slice(0,25).join('\n'):'No eligible notifications under the current routes.';Logger.log(msg);SpreadsheetApp.getUi().alert('Dry-run preview',msg,SpreadsheetApp.getUi().ButtonSet.OK);}
function pollCareerAlerts(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
  try{
    const c=getSettings_();if(c.MODE!=='LIVE'||PropertiesService.getScriptProperties().getProperty('CAREER_LIVE_ARMED')!=='YES')return;
    configValid_(c);credentials_();const now=new Date();const s=book_().getSheetByName('Notifications');const idx=notificationIndex_();
    const counts={SMS:countAttempts_(idx,'SMS',now),CALL:countAttempts_(idx,'CALL',now)};
    let attempted=0;for(const e of eventList_())for(const ch of channels_(c['ROUTE_'+e.type])){
      let item=idx.get(e.id+'|'+ch);if(item&&item.values[4]!=='PENDING')continue;
      if(!item){const row=appendNotification_(s,e,ch,'PENDING',now,'Queued');item={row,values:[e.id,ch,e.type,e.app,'PENDING']};idx.set(e.id+'|'+ch,item);}
      if(inQuiet_(localClock_(now),c.QUIET_START,c.QUIET_END)||counts[ch]>=Number(c[ch==='SMS'?'MAX_SMS_PER_DAY':'MAX_CALLS_PER_DAY'])||attempted>=10)continue;
      // Recheck kill switch before each paid request.
      if(PropertiesService.getScriptProperties().getProperty('CAREER_LIVE_ARMED')!=='YES'||getSettings_().MODE!=='LIVE')return;
      attempt_(s,item.row,e,ch);counts[ch]++;attempted++;
    }
    refreshStatuses_(s,idx);
  }catch(err){log_('ERROR','Poll',String(err.message||'Unknown error').slice(0,200));throw err;}finally{lock.releaseLock();}
}
function refreshStatuses_(s,idx){const c=credentials_();let n=0;for(const {row,values:r} of idx.values()){
  if(n>=15)break;if(r[4]!=='ACCEPTED'||!r[7]||['delivered','undelivered','failed','completed','busy','no-answer','canceled','sent'].includes(String(r[8])))continue;
  const prefix=r[1]==='SMS'?'Messages/':'Calls/';if(!/^(SM|MM|CA)[0-9a-fA-F]{32}$/.test(r[7]))continue;
  try{const result=twilioRequest_(c,prefix+r[7]+'.json','get');if(result.sid){s.getRange(row,9).setValue(result.status);s.getRange(row,11).setValue(new Date());}}catch(_){}n++;
}}
function manualTest_(ch){const c=getSettings_();configValid_(c);credentials_();const ui=SpreadsheetApp.getUi();if(ui.alert('Send paid test?', 'Send one '+ch+' to the number stored privately in ALERT_TO_NUMBER? Charges or trial units may apply.',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  const lock=LockService.getScriptLock();lock.waitLock(30000);try{const now=new Date();const idx=notificationIndex_();if(inQuiet_(localClock_(now),c.QUIET_START,c.QUIET_END))throw Error('Test blocked during quiet hours.');if(countAttempts_(idx,ch,now)>=Number(c[ch==='SMS'?'MAX_SMS_PER_DAY':'MAX_CALLS_PER_DAY']))throw Error('Daily limit reached.');const s=book_().getSheetByName('Notifications');const e={id:'TEST:'+Utilities.getUuid(),type:'TEST',app:''};const row=appendNotification_(s,e,ch,'PENDING',now,'Manually approved test.');const result=attempt_(s,row,e,ch);Logger.log(result.state+' '+result.status);}finally{lock.releaseLock();}}
function testCareerSms(){manualTest_('SMS');}
function testCareerCall(){manualTest_('CALL');}
