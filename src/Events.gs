function rows_(name,expected){const s=book_().getSheetByName(name);if(!s)throw Error('Missing '+name);const values=s.getDataRange().getValues();const headers=values[0].map(String);for(const [col,key] of expected)if(headers[col]!==key)throw Error('Unexpected '+name+' header '+key);return values.slice(1);}
function responseType_(type){return {'Acknowledgment only':'ACKNOWLEDGMENT','Recruiter reply':'EMPLOYER_REPLY','Assessment':'ASSESSMENT','Interview invitation':'INTERVIEW','Offer':'OFFER','Rejection':'REJECTION','Other substantive reply':'EMPLOYER_REPLY','Progression':'PROGRESSION'}[type]||null;}
function eventList_(){
  const a=rows_('Applications',[[0,'Application ID'],[10,'Status'],[11,'Date submitted'],[13,'Confirmation / proof']]);
  const r=rows_('Responses',[[0,'Response ID'],[1,'Application ID'],[4,'Response type']]);
  const q=rows_('Approval Queue',[[0,'Approval ID'],[1,'Application ID'],[6,'Status']]);
  const out=[];const apps=new Set(a.filter(x=>x[0]).map(x=>String(x[0])));
  a.forEach(x=>{if(x[0]&&x[11] instanceof Date&&!isNaN(x[11])&&x[13])out.push({id:'APP:'+x[0]+':SUBMITTED',type:'SUBMITTED',app:String(x[0])});});
  r.forEach(x=>{const t=responseType_(String(x[4]));if(x[0]&&x[1]&&apps.has(String(x[1]))&&t)out.push({id:'RESPONSE:'+x[0],type:t,app:String(x[1])});});
  q.forEach(x=>{if(x[0]&&x[6]==='Pending')out.push({id:'APPROVAL:'+x[0],type:'APPROVAL',app:String(x[1]||'')});});
  const seen=new Set();return out.filter(e=>{if(seen.has(e.id))return false;seen.add(e.id);return true;});
}
function notificationText_(type){return 'Boss, '+(CA.labels[type]||'A job search update arrived.')+' Please check your job search dashboard.';}
