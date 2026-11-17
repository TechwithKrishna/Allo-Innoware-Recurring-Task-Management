/**************** ALLO INNOWARE — Recurring Task Automation ****************
 * Import the workbook into Google Sheets, then paste this entire file into
 * Extensions > Apps Script as Code.gs. Run setupPrototype() once.
 * Configure MANAGER_EMAIL and LEADERSHIP_EMAIL in the Config sheet first.
 ****************************************************************************/

const SHEETS = {
  MASTER: 'Task Master', QUEUE: 'Employee Work Queue',
  LOG: 'Cycle Log', DASH: 'Manager Dashboard', ALERTS: 'Alert Log',
  CONFIG: 'Config'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Allo Ops')
    .addItem('Setup / Repair Automation', 'setupPrototype')
    .addItem('Refresh Employee Queue + Dashboard', 'refreshViews')
    .addItem('Run Due Alerts Now (test)', 'processDueTasks')
    .addToUi();
}

function setupPrototype() {
  const ss = SpreadsheetApp.getActive();
  Object.values(SHEETS).forEach(n => { if (!ss.getSheetByName(n)) throw new Error('Missing sheet: '+n); });
  ensureInitialCycles_();
  refreshViews();
  // Remove duplicate project triggers before installing the two required triggers.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['processDueTasks','refreshViews'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processDueTasks').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('refreshViews').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Setup complete. Hourly alert and view-refresh triggers installed.');
}

function ensureInitialCycles_() {
  const ss=SpreadsheetApp.getActive(), master=ss.getSheetByName(SHEETS.MASTER), log=ss.getSheetByName(SHEETS.LOG);
  const m=master.getDataRange().getValues(), l=log.getDataRange().getValues();
  const existing=new Set(l.slice(1).map(r=>String(r[0])));
  const now=new Date();
  for(let i=1;i<m.length;i++){
    const [taskId,name,owner,freq,days,proof,dueTime,active]=m[i];
    if(!taskId || active!==true) continue;
    const hasOpen=l.slice(1).some(r=>String(r[1])===String(taskId) && !['COMPLETED','CANCELLED'].includes(String(r[7])));
    if(hasOpen) continue;
    const id=String(taskId)+'-'+Utilities.formatDate(now, getTz_(), 'yyyyMMdd')+'-001';
    if(existing.has(id)) continue;
    const due=initialDueForTask_(Number(days), String(freq), now);
    log.appendRow([id,taskId,name,owner,now,due,dueTime,'PENDING',proof,'','','',false,'','','','']);
  }
}

function initialDueForTask_(intervalDays, frequency, now) {
  const d=new Date(now);
  if(String(frequency).toLowerCase()==='daily') {
    while(d.getDay()===0 || d.getDay()===6) d.setDate(d.getDate()+1);
    return new Date(d.getFullYear(),d.getMonth(),d.getDate());
  }
  d.setDate(d.getDate()+intervalDays);
  return new Date(d.getFullYear(),d.getMonth(),d.getDate());
}

function nextDueForTask_(completedAt, intervalDays, frequency, fallbackNow) {
  const base=completedAt || fallbackNow || new Date();
  const d=new Date(base);
  if(String(frequency).toLowerCase()==='daily') {
    d.setDate(d.getDate()+1);
    while(d.getDay()===0 || d.getDay()===6) d.setDate(d.getDate()+1);
    return new Date(d.getFullYear(),d.getMonth(),d.getDate());
  }
  d.setDate(d.getDate()+intervalDays);
  return new Date(d.getFullYear(),d.getMonth(),d.getDate());
}

function onEdit(e) {
  if(!e || !e.range) return;
  const sh=e.range.getSheet(), name=sh.getName(), row=e.range.getRow(), col=e.range.getColumn();
  if(row<2) return;
  if(name===SHEETS.QUEUE && col===10 && String(e.value).toUpperCase()==='TRUE') {
    completeCycleFromQueue_(row);
  } else if(name===SHEETS.DASH && col===9 && String(e.value).toUpperCase()==='TRUE') {
    acknowledgeFromDashboard_(row);
  }
}

function completeCycleFromQueue_(row) {
  const ss=SpreadsheetApp.getActive(), q=ss.getSheetByName(SHEETS.QUEUE), log=ss.getSheetByName(SHEETS.LOG);
  const v=q.getRange(row,1,1,12).getValues()[0];
  const [cycleId,taskId,task,owner,dueDate,dueTime,status,proofType,proof,done]=v;
  if(!cycleId) return;
  if(!proof || String(proof).trim()==='') {
    q.getRange(row,10).setValue(false);
    SpreadsheetApp.getActive().toast('Enter the required Ticket ID or daily update reference before marking complete.');
    return;
  }
  const data=log.getDataRange().getValues();
  const idx=data.findIndex((r,i)=>i>0 && String(r[0])===String(cycleId));
  if(idx<1 || String(data[idx][7])==='COMPLETED') { q.getRange(row,10).setValue(false); return; }
  const actor=Session.getActiveUser().getEmail() || String(owner);
  const completedAt=new Date();
  log.getRange(idx+1,8,1,5).setValues([['COMPLETED',String(data[idx][8]),String(proof),actor,completedAt]]);
  q.getRange(row,11,1,2).setValues([[actor,completedAt]]);
  createNextCycle_(taskId, completedAt);
  refreshViews();
}

function createNextCycle_(taskId, completedAt) {
  const ss=SpreadsheetApp.getActive(), master=ss.getSheetByName(SHEETS.MASTER), log=ss.getSheetByName(SHEETS.LOG);
  const m=master.getDataRange().getValues(), l=log.getDataRange().getValues();
  const task=m.slice(1).find(r=>String(r[0])===String(taskId));
  if(!task || task[7]!==true) return;
  // Idempotency: do not create another open cycle for this task.
  if(l.slice(1).some(r=>String(r[1])===String(taskId) && !['COMPLETED','CANCELLED'].includes(String(r[7])))) return;
  const [id,name,owner,freq,days,proof,dueTime]=task;
  const due=nextDueForTask_(completedAt,Number(days),String(freq),completedAt);
  const cycleId=String(id)+'-'+Utilities.formatDate(completedAt,getTz_(),'yyyyMMddHHmmss');
  log.appendRow([cycleId,id,name,owner,completedAt,due,dueTime,'PENDING',proof,'','','',false,'','','','']);
}

function refreshViews() {
  const ss=SpreadsheetApp.getActive(), log=ss.getSheetByName(SHEETS.LOG);
  const q=ss.getSheetByName(SHEETS.QUEUE), d=ss.getSheetByName(SHEETS.DASH);
  const data=log.getDataRange().getValues();
  // Preserve employee-entered proof for each still-open cycle across refreshes.
  const priorQueue=new Map();
  if(q.getLastRow()>1) q.getRange(2,1,q.getLastRow()-1,12).getValues().forEach(r=>{
    if(r[0]) priorQueue.set(String(r[0]), {proof:r[8], checked:r[9]===true});
  });
  if(q.getLastRow()>1) q.getRange(2,1,q.getLastRow()-1,12).clearContent();
  if(d.getLastRow()>1) d.getRange(2,1,d.getLastRow()-1,12).clearContent();
  const queueRows=[], dashRows=[];
  for(let i=1;i<data.length;i++){
    const r=data[i], status=String(r[7]);
    if(['COMPLETED','CANCELLED'].includes(status)) continue;
    const prior=priorQueue.get(String(r[0])) || {proof:'',checked:false};
    queueRows.push([r[0],r[1],r[2],r[3],r[5],r[6],status,r[8],prior.proof,prior.checked,'','']);
    const due=combineDateTime_(r[5],r[6]), hours=Math.max(0,(new Date()-due)/3600000);
    dashRows.push([r[0],r[2],r[3],r[5],r[6],status,Math.round(hours*10)/10,r[9],r[12]===true,r[13],r[14],r[16]]);
  }
  if(queueRows.length) q.getRange(2,1,queueRows.length,12).setValues(queueRows);
  if(dashRows.length) d.getRange(2,1,dashRows.length,12).setValues(dashRows);
  // Restore checkbox validation for employee completion and manager acknowledgement.
  if(queueRows.length) q.getRange(2,10,queueRows.length,1).insertCheckboxes();
  if(dashRows.length) d.getRange(2,9,dashRows.length,1).insertCheckboxes();
}

function processDueTasks() {
  const ss=SpreadsheetApp.getActive(), log=ss.getSheetByName(SHEETS.LOG), alerts=ss.getSheetByName(SHEETS.ALERTS);
  const rows=log.getDataRange().getValues(), cfg=getConfig_(), now=new Date();
  const sent=new Set(alerts.getDataRange().getValues().slice(1).map(r=>String(r[0])));
  for(let i=1;i<rows.length;i++){
    const r=rows[i], cycleId=String(r[0]), status=String(r[7]);
    if(['COMPLETED','CANCELLED'].includes(status)) continue;
    const due=combineDateTime_(r[5],r[6]), managerAlertAt=new Date(due.getTime()+Number(cfg.MANAGER_ALERT_GRACE_MINUTES||15)*60000);
    const managerEmail=cfg.MANAGER_EMAIL || '';
    const managerKey=cycleId+'|MANAGER';
    if(now>=managerAlertAt && managerEmail && !sent.has(managerKey)){
      const subject=(cfg.EMAIL_SUBJECT_PREFIX||'[Allo Ops] ')+'Pending task: '+r[2];
      const body='Task is still pending.\\nEmployee: '+r[3]+'\\nTask: '+r[2]+'\\nDue: '+due+'\\nCycle: '+cycleId+'\\nPlease review the Manager Dashboard and acknowledge the issue.';
      sendAndLog_(alerts,managerKey,cycleId,'MANAGER',managerEmail,subject,body);
      log.getRange(i+1,16).setValue(now);
      sent.add(managerKey);
    }
    const alertSent=r[15] ? new Date(r[15]) : managerAlertAt;
    const ackAt=r[14] ? new Date(r[14]) : null;
    const escalationByDue=new Date(due.getTime()+Number(cfg.LEADERSHIP_ESCALATION_HOURS||24)*3600000);
    const escalationByAck=addBusinessHours_(alertSent,Number(cfg.MANAGER_ACK_SLA_BUSINESS_HOURS||4));
    const acknowledgedOnTime=ackAt && ackAt.getTime()<=escalationByAck.getTime();
    const escalationAt=acknowledgedOnTime ? escalationByDue : new Date(Math.min(escalationByDue.getTime(),escalationByAck.getTime()));
    const leadershipEmail=cfg.LEADERSHIP_EMAIL || '';
    const leadershipKey=cycleId+'|LEADERSHIP';
    if(now>=escalationAt && leadershipEmail && !sent.has(leadershipKey)){
      const reason=acknowledgedOnTime ? 'Task remains unresolved beyond the 24-hour SLA.' : 'Manager acknowledgement SLA was missed or task remains unresolved beyond 24 hours.';
      const subject=(cfg.EMAIL_SUBJECT_PREFIX||'[Allo Ops] ')+'ESCALATION: '+r[2];
      const body=reason+'\\nEmployee: '+r[3]+'\\nTask: '+r[2]+'\\nDue: '+due+'\\nManager acknowledged: '+(ackAt||'No')+'\\nCycle: '+cycleId;
      sendAndLog_(alerts,leadershipKey,cycleId,'LEADERSHIP',leadershipEmail,subject,body);
      log.getRange(i+1,17).setValue(now);
      sent.add(leadershipKey);
    }
  }
  refreshViews();
}

function acknowledgeFromDashboard_(row) {
  const ss=SpreadsheetApp.getActive(), d=ss.getSheetByName(SHEETS.DASH), log=ss.getSheetByName(SHEETS.LOG);
  const cycleId=String(d.getRange(row,1).getValue());
  if(!cycleId) return;
  const data=log.getDataRange().getValues(), idx=data.findIndex((r,i)=>i>0 && String(r[0])===cycleId);
  if(idx<1) return;
  const actor=Session.getActiveUser().getEmail() || 'Manager';
  log.getRange(idx+1,13,1,3).setValues([[true,actor,new Date()]]);
  refreshViews();
}

function sendAndLog_(sheet,key,cycleId,type,to,subject,body) {
  try {
    MailApp.sendEmail(to,subject,body);
    sheet.appendRow([key,cycleId,type,to,new Date(),'SENT']);
  } catch(err) {
    sheet.appendRow([key,cycleId,type,to,new Date(),'FAILED: '+err.message]);
  }
}

function getConfig_() {
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG), vals=sh.getDataRange().getValues(), o={};
  vals.slice(1).forEach(r=>{if(r[0]) o[String(r[0])]=r[1];});
  return o;
}
function getTz_(){ return getConfig_().TIMEZONE || Session.getScriptTimeZone() || 'Asia/Kolkata'; }
function combineDateTime_(date,time) {
  const d=new Date(date), parts=String(time||'17:00').split(':');
  d.setHours(Number(parts[0]||17),Number(parts[1]||0),0,0); return d;
}
function addBusinessHours_(start,hours) {
  let d=new Date(start), remaining=hours;
  while(remaining>0) {
    d.setHours(d.getHours()+1);
    if(d.getDay()!==0 && d.getDay()!==6 && d.getHours()>=9 && d.getHours()<18) remaining--;
    if(d.getHours()>=18) d.setDate(d.getDate()+1),d.setHours(9,0,0,0);
    if(d.getDay()===0) d.setDate(d.getDate()+1),d.setHours(9,0,0,0);
    if(d.getDay()===6) d.setDate(d.getDate()+2),d.setHours(9,0,0,0);
  }
  return d;
}
