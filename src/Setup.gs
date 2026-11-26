function onOpen(){SpreadsheetApp.getUi().createMenu('Career Assistant')
  .addItem('Initialize tracker tabs','setupCareerAssistant')
  .addItem('Preview notifications (no sends)','previewCareerAlerts')
  .addSeparator().addItem('Send one test SMS','testCareerSms').addItem('Make one test call','testCareerCall')
  .addSeparator().addItem('Arm live notifications','armCareerAlerts').addItem('Disable live notifications','disableCareerAlerts')
  .addItem('Install / replace polling trigger','installCareerTrigger').addItem('Remove polling trigger','removeCareerTrigger').addToUi();}
function setupCareerAssistant(){
  const b=SpreadsheetApp.getActiveSpreadsheet();if(!b)throw Error('Open the tracker and run from its bound Apps Script.');
  ['Applications','Responses','Interviews'].forEach(n=>{if(!b.getSheetByName(n))throw Error('Expected tracker tab missing: '+n);});
  PropertiesService.getScriptProperties().setProperty('CAREER_SHEET_ID',b.getId());
  Object.entries(CA.sheets).forEach(([name,headers])=>{
    let s=b.getSheetByName(name); if(!s)s=b.insertSheet(name);
    if(s.getLastRow()===0){s.getRange(1,1,1,headers.length).setValues([headers]);s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e3e8ed').setWrap(true);s.setFrozenRows(1);s.setColumnWidths(1,headers.length,160);if(name==='Settings')s.setColumnWidth(3,520);}
    else {const actual=s.getRange(1,1,1,headers.length).getDisplayValues()[0];if(actual.join('|')!==headers.join('|'))throw Error('Existing '+name+' headers differ; no overwrite performed.');}
  });
  // Extend existing tracker response dropdowns with an explicitly verified progression event.
  [['Responses','E'],['Applications','P']].forEach(([name,col])=>{
    const sh=b.getSheetByName(name), rule=sh.getRange(col+'2').getDataValidation();
    if(rule&&rule.getCriteriaType()===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST){
      const args=rule.getCriteriaValues(), values=args[0];
      if(!values.includes('Progression'))sh.getRange(col+'2:'+col+sh.getMaxRows()).setDataValidation(rule.copy().requireValueInList([...values,'Progression'],args[1]!==false).build());
    }
  });
  const s=b.getSheetByName('Settings');const existing=new Set(s.getDataRange().getDisplayValues().slice(1).map(r=>r[0]));
  const add=settingsDefaults_().filter(r=>!existing.has(r[0]));if(add.length)s.getRange(s.getLastRow()+1,1,add.length,3).setValues(add);
  s.getRange(2,2,s.getLastRow()-1,1).setNumberFormat('@');
  const rows=s.getDataRange().getDisplayValues();rows.forEach((r,i)=>{if(r[0].startsWith('ROUTE_'))s.getRange(i+1,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OFF','SMS','CALL','BOTH'],true).setAllowInvalid(false).build());});
  b.getSheetByName('Notifications').getRange('F2:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');b.getSheetByName('Notifications').getRange('K2:K').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  b.getSheetByName('Approval Queue').getRange('C2:C').setNumberFormat('yyyy-mm-dd hh:mm:ss');b.getSheetByName('Approval Queue').getRange('F2:F').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  b.getSheetByName('System Logs').getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  log_('INFO','Setup','Tracker tabs initialized. Live alerts are not armed by setup.');
}
function installCareerTrigger(){const c=getSettings_();const n=Number(c.POLL_MINUTES);if(![1,5,10,15,30].includes(n))throw Error('POLL_MINUTES must be 1, 5, 10, 15 or 30.');removeCareerTrigger();ScriptApp.newTrigger('pollCareerAlerts').timeBased().everyMinutes(n).create();log_('INFO','Trigger','Polling trigger installed.');}
function removeCareerTrigger(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='pollCareerAlerts').forEach(t=>ScriptApp.deleteTrigger(t));}
function disableCareerAlerts(){PropertiesService.getScriptProperties().deleteProperty('CAREER_LIVE_ARMED');setSetting_('MODE','DRY_RUN');log_('INFO','Mode','Live notifications disabled.');}
function armCareerAlerts(){
  const c=getSettings_();configValid_(c);credentials_();if(!CA.routes.some(k=>channels_(c['ROUTE_'+k]).length))throw Error('Choose at least one notification route first.');
  const ui=SpreadsheetApp.getUi();if(ui.alert('Enable paid alerts?', 'This will skip all currently logged events and enable future SMS/calls to your verified personal number using your chosen policy. Continue?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  const lock=LockService.getScriptLock();lock.waitLock(30000);try{
    const s=book_().getSheetByName('Notifications');const old=notificationIndex_();const now=new Date();
    eventList_().forEach(e=>['SMS','CALL'].forEach(ch=>{const item=old.get(e.id+'|'+ch);if(!item)appendNotification_(s,e,ch,'BASELINE',now,'Existing event skipped when arming.');else if(item.values[4]==='PENDING'){s.getRange(item.row,5).setValue('BASELINE');s.getRange(item.row,10,1,2).setValues([['Existing pending event skipped when arming.',now]]);}}));
    setSetting_('MODE','LIVE');PropertiesService.getScriptProperties().setProperty('CAREER_LIVE_ARMED','YES');log_('INFO','Mode','Live alerts armed.');
  }finally{lock.releaseLock();}
}
