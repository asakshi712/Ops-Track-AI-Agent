/**
 * OpsTrack - Recurring Task Tracking, Pending Reports & Escalation Automation
 * Assignment prototype for Allo Innoware.
 *
 * Setup:
 * 1. Create a new Google Sheet.
 * 2. Extensions > Apps Script. Replace Code.gs with this file.
 * 3. Add an HTML file named "Sidebar" and paste Sidebar.html.
 * 4. Save, reload the spreadsheet, then OpsTrack > Initialize / Refresh.
 * 5. Set manager and leadership email addresses in Config.
 * 6. OpsTrack > Install daily automation. Authorize the requested permissions.
 *
 * This is a prototype. Use dummy data until the business approves production use.
 */

const SHEETS = {
  master: 'Task_Master',
  console: 'Employee_Console',
  log: 'Completion_Log',
  alerts: 'Alerts_Log',
  config: 'Config'
};

const MASTER_HEADERS = [
  'Task ID', 'Task Name', 'Assigned To', 'Frequency Days', 'Next Due Date',
  'Required Proof', 'Active', 'Cycle Number', 'Last Completed At',
  'Last Completed By', 'Last Proof'
];
const CONSOLE_HEADERS = [
  'Task ID', 'Task Name', 'Assigned To', 'Frequency Days', 'Due Date',
  'Status', 'Required Proof', 'Ticket ID / Daily Proof', 'Mark Complete',
  'Completed By', 'Completed At', 'Manager Acknowledged', 'Manager Notes',
  'Cycle Number'
];
const LOG_HEADERS = [
  'Event ID', 'Timestamp', 'Task ID', 'Task Name', 'Cycle Number',
  'Assigned To', 'Completed By', 'Completion Timestamp', 'Proof Type',
  'Proof / Ticket ID', 'Result'
];
const ALERT_HEADERS = [
  'Alert Key', 'Created At', 'Task ID', 'Task Name', 'Cycle Number',
  'Assigned To', 'Alert Type', 'Recipient', 'Sent At', 'Manager Acknowledged At',
  'Resolved At', 'Details'
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('OpsTrack')
    .addItem('Initialize / Refresh system', 'initializeOpsTrack')
    .addItem('Refresh employee console', 'refreshEmployeeConsole')
    .addItem('Run SLA check now', 'runSlaCheck')
    .addItem('Install daily automation', 'installDailyAutomation')
    .addItem('Open quick guide', 'showQuickGuide')
    .addToUi();
}

function initializeOpsTrack() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEETS.master, MASTER_HEADERS);
  ensureSheet_(ss, SHEETS.console, CONSOLE_HEADERS);
  ensureSheet_(ss, SHEETS.log, LOG_HEADERS);
  ensureSheet_(ss, SHEETS.alerts, ALERT_HEADERS);
  setupConfig_(ss);
  seedTasksIfEmpty_(ss);
  formatSheets_(ss);
  refreshEmployeeConsole();
  SpreadsheetApp.getUi().alert(
    'OpsTrack initialized. Update Config with real recipients, review the dummy tasks, then install daily automation.'
  );
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#17324D').setFontColor('#FFFFFF');
  return sh;
}

function setupConfig_(ss) {
  const sh = ss.getSheetByName(SHEETS.config);
  sh.clear();
  sh.getRange(1,1,1,3).setValues([['Setting','Value','Purpose']]);
  sh.getRange(2,1,7,3).setValues([
    ['Manager Email','replace-with-manager@example.com','Recipient for first overdue alert'],
    ['Leadership Email','replace-with-leadership@example.com','Recipient for escalation email'],
    ['Manager Alert Hour (0-23)',9,'Daily trigger runs near this hour; Google may vary trigger time'],
    ['Escalation Delay Hours',24,'Escalate unresolved issue after this many hours from first manager alert'],
    ['Daily Task Due Hour (0-23)',17,'Daily task is treated as overdue after this local hour'],
    ['Timezone',Session.getScriptTimeZone(),'Set project timezone in Apps Script project settings'],
    ['Test Mode','YES','YES labels emails as TEST and uses dummy tasks; set NO only after review']
  ]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#17324D').setFontColor('#FFFFFF');
  sh.autoResizeColumns(1,3);
}

function seedTasksIfEmpty_(ss) {
  const sh = ss.getSheetByName(SHEETS.master);
  if (sh.getLastRow() > 1) return;
  const today = new Date();
  const rows = [
    ['T-001','Daily Operations Update','Employee A',1,atHour_(today,17),'WhatsApp update link or short evidence',true,1,'','',''],
    ['T-002','Portal Task A','Employee B',7,addDays_(today,-1),'Ticket ID',true,1,'','',''],
    ['T-003','Portal Task B','Employee C',15,addDays_(today,2),'Ticket ID',true,1,'','',''],
    ['T-004','Portal Task C','Employee D',30,addDays_(today,5),'Ticket ID',true,1,'','','']
  ];
  sh.getRange(2,1,rows.length,MASTER_HEADERS.length).setValues(rows);
  sh.getRange(2,5,rows.length,1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function refreshEmployeeConsole() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SHEETS.master);
  const out = ss.getSheetByName(SHEETS.console);
  if (!master || !out) throw new Error('Run Initialize / Refresh system first.');
  const values = master.getDataRange().getValues();
  const rows = [];
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0] || r[6] !== true) continue;
    const due = r[4] instanceof Date ? r[4] : new Date(r[4]);
    const cycle = Number(r[7] || 1);
    const status = due.getTime() < now.getTime() ? 'OVERDUE' : 'PENDING';
    rows.push([
      r[0], r[1], r[2], r[3], due, status, r[5],
      '', false, '', '', false, '', cycle
    ]);
  }

  // Preserve in-progress proof, completion and manager review fields for the same task/cycle.
  const old = out.getDataRange().getValues();
  const oldMap = {};
  for (let i = 1; i < old.length; i++) {
    const key = String(old[i][0]) + '|' + String(old[i][13]);
    oldMap[key] = old[i];
  }
  const merged = rows.map(r => {
    const oldRow = oldMap[String(r[0]) + '|' + String(r[13])];
    if (oldRow) {
      r[7] = oldRow[7] || '';
      r[8] = oldRow[8] === true;
      r[9] = oldRow[9] || '';
      r[10] = oldRow[10] || '';
      r[11] = oldRow[11] === true;
      r[12] = oldRow[12] || '';
    }
    return r;
  });

  if (out.getLastRow() > 1) out.getRange(2,1,out.getLastRow()-1,CONSOLE_HEADERS.length).clearContent();
  if (merged.length) {
    out.getRange(2,1,merged.length,CONSOLE_HEADERS.length).setValues(merged);
    out.getRange(2,5,merged.length,1).setNumberFormat('yyyy-mm-dd hh:mm');
    out.getRange(2,9,merged.length,1).insertCheckboxes();
    out.getRange(2,12,merged.length,1).insertCheckboxes();
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDING','OVERDUE','COMPLETED'], true).build();
    out.getRange(2,6,merged.length,1).setDataValidation(rule);
  }
  formatSheets_(ss);
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.console || e.range.getRow() < 2) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const ss = e.source;
  const taskId = String(sh.getRange(row,1).getValue());
  const cycle = Number(sh.getRange(row,14).getValue());
  if (!taskId) return;

  // Column 12: manager acknowledgement. Capture acknowledgement timestamp in alert log.
  if (col === 12 && e.range.getValue() === true) {
    recordManagerAcknowledgement_(taskId, cycle, String(sh.getRange(row,13).getValue() || ''));
    return;
  }

  // Column 9: employee marks complete. Completion is accepted only with required proof.
  if (col !== 9 || e.range.getValue() !== true) return;
  const proof = String(sh.getRange(row,8).getValue() || '').trim();
  const required = String(sh.getRange(row,7).getValue() || '').toLowerCase();
  if (!proof) {
    e.range.setValue(false);
    SpreadsheetApp.getActive().toast('Add Ticket ID or daily-task proof before marking complete.', 'Proof required', 5);
    return;
  }
  const completedBy = Session.getActiveUser().getEmail() || String(sh.getRange(row,3).getValue());
  const completedAt = new Date();
  const master = ss.getSheetByName(SHEETS.master);
  const mdata = master.getDataRange().getValues();
  let masterRow = -1;
  for (let i=1; i<mdata.length; i++) {
    if (String(mdata[i][0]) === taskId && Number(mdata[i][7] || 1) === cycle) { masterRow = i+1; break; }
  }
  if (masterRow < 0) {
    e.range.setValue(false);
    SpreadsheetApp.getActive().toast('Task cycle could not be found. Refresh and retry.', 'Cycle mismatch', 5);
    return;
  }

  // Idempotency guard: do not log a second completion for the same task cycle.
  const log = ss.getSheetByName(SHEETS.log);
  const logs = log.getDataRange().getValues();
  const eventId = taskId + '-C' + cycle;
  for (let i=1; i<logs.length; i++) {
    if (String(logs[i][0]) === eventId && String(logs[i][10]) === 'COMPLETED') {
      e.range.setValue(false);
      SpreadsheetApp.getActive().toast('This task cycle is already logged as completed.', 'Duplicate prevented', 5);
      refreshEmployeeConsole();
      return;
    }
  }

  const proofType = required.indexOf('ticket') >= 0 ? 'Ticket ID' : 'Daily / other proof';
  log.appendRow([eventId, new Date(), taskId, sh.getRange(row,2).getValue(), cycle,
    sh.getRange(row,3).getValue(), completedBy, completedAt, proofType, proof, 'COMPLETED']);

  const frequency = Number(master.getRange(masterRow,4).getValue());
  const nextDue = addDays_(completedAt, frequency);
  master.getRange(masterRow,5).setValue(nextDue);
  master.getRange(masterRow,8).setValue(cycle + 1);
  master.getRange(masterRow,9).setValue(completedAt);
  master.getRange(masterRow,10).setValue(completedBy);
  master.getRange(masterRow,11).setValue(proof);
  closeOpenAlerts_(taskId, cycle);
  refreshEmployeeConsole();
  SpreadsheetApp.getActive().toast('Completion logged. Next cycle due ' + nextDue.toLocaleString(), 'OpsTrack', 6);
}

function runSlaCheck() {
  const ss = SpreadsheetApp.getActive();
  const cfg = getConfig_();
  const master = ss.getSheetByName(SHEETS.master);
  const consoleSh = ss.getSheetByName(SHEETS.console);
  if (!master || !consoleSh) throw new Error('Initialize the system first.');
  refreshEmployeeConsole();
  const data = consoleSh.getDataRange().getValues();
  const now = new Date();
  const manager = String(cfg['Manager Email'] || '').trim();
  const leadership = String(cfg['Leadership Email'] || '').trim();
  const delayHours = Number(cfg['Escalation Delay Hours'] || 24);
  const testMode = String(cfg['Test Mode'] || 'YES').toUpperCase() === 'YES';

  for (let i=1; i<data.length; i++) {
    const r = data[i];
    if (!r[0] || r[5] === 'COMPLETED') continue;
    const due = r[4] instanceof Date ? r[4] : new Date(r[4]);
    if (isNaN(due.getTime()) || now <= due) continue;
    const taskId = String(r[0]), taskName = String(r[1]), cycle = Number(r[13]);
    const key = taskId + '|C' + cycle;
    let firstAlert = findAlert_(key, 'MANAGER');
    if (!firstAlert && manager && manager.indexOf('@') > 0) {
      const subject = (testMode ? '[TEST] ' : '') + 'OpsTrack overdue task: ' + taskName;
      const body = [
        'An operational task is overdue.',
        '',
        'Task: ' + taskName,
        'Task ID / cycle: ' + taskId + ' / ' + cycle,
        'Assigned to: ' + r[2],
        'Due: ' + due,
        'Required proof: ' + r[6],
        '',
        'Manager action: review the Employee_Console row, follow up with the owner, then tick Manager Acknowledged and add a short note.',
        'The system checks for acknowledgement and resolution during its next scheduled run.'
      ].join('\n');
      MailApp.sendEmail(manager, subject, body);
      firstAlert = logAlert_(key, r, 'MANAGER', manager, 'Manager email sent; awaiting acknowledgement/resolution.');
    }

    if (firstAlert) {
      const sentAt = firstAlert.sentAt || firstAlert.createdAt;
      const acknowledged = !!firstAlert.ackAt;
      const ageHours = (now.getTime() - new Date(sentAt).getTime()) / 3600000;
      if (ageHours >= delayHours && !findAlert_(key, 'LEADERSHIP')) {
        if (leadership && leadership.indexOf('@') > 0) {
          const subject = (testMode ? '[TEST] ' : '') + 'OpsTrack escalation: unresolved task ' + taskName;
          const body = [
            'Leadership escalation: a task remains unresolved beyond the configured SLA.',
            '',
            'Task: ' + taskName,
            'Task ID / cycle: ' + taskId + ' / ' + cycle,
            'Assigned to: ' + r[2],
            'Due: ' + due,
            'Manager acknowledgement recorded: ' + (acknowledged ? 'Yes' : 'No'),
            'Elapsed since manager alert: ' + ageHours.toFixed(1) + ' hours',
            '',
            'Please review the task and coordinate resolution.'
          ].join('\n');
          MailApp.sendEmail(leadership, subject, body);
          logAlert_(key, r, 'LEADERSHIP', leadership, 'Escalated after SLA delay.');
        }
      }
    }
  }
}

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.config);
  const vals = sh.getDataRange().getValues();
  const out = {};
  for (let i=1; i<vals.length; i++) if (vals[i][0]) out[String(vals[i][0])] = vals[i][1];
  return out;
}

function findAlert_(key, type) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.alerts);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === key && String(data[i][6]) === type) {
      return { row: i+1, createdAt: data[i][1], sentAt: data[i][8] || data[i][1], ackAt: data[i][9], resolvedAt: data[i][10] };
    }
  }
  return null;
}

function logAlert_(key, r, type, recipient, details) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.alerts);
  const now = new Date();
  sh.appendRow([key, now, r[0], r[1], r[13], r[2], type, recipient, now, '', '', details]);
  return { createdAt: now, sentAt: now, ackAt: '', resolvedAt: '' };
}

function recordManagerAcknowledgement_(taskId, cycle, note) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.alerts);
  const data = sh.getDataRange().getValues();
  const key = taskId + '|C' + cycle;
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === key && String(data[i][6]) === 'MANAGER' && !data[i][9]) {
      sh.getRange(i+1,10).setValue(new Date());
      sh.getRange(i+1,12).setValue('Manager acknowledged. Note: ' + note);
    }
  }
}

function closeOpenAlerts_(taskId, cycle) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.alerts);
  const data = sh.getDataRange().getValues();
  const key = taskId + '|C' + cycle;
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === key && !data[i][10]) sh.getRange(i+1,11).setValue(new Date());
  }
}

function installDailyAutomation() {
  // Remove duplicate triggers for this function before installing one daily trigger.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runSlaCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runSlaCheck').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Hourly SLA check trigger installed. Google may run time-based triggers with some timing variation.');
}

function showQuickGuide() {
  SpreadsheetApp.getUi().alert(
    'Employee: use Employee_Console only. Enter Ticket ID or daily proof in column H, then tick column I.\\n\\n' +
    'Manager: review overdue rows, tick column L, and enter a note in column M.\\n\\n' +
    'Admin: configure recipient emails in Config, run OpsTrack > Run SLA check now to test, and inspect Completion_Log / Alerts_Log.'
  );
}

function formatSheets_(ss) {
  Object.values(SHEETS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setFrozenRows(1);
    sh.getDataRange().setVerticalAlignment('middle');
    sh.autoResizeColumns(1, Math.min(sh.getLastColumn() || 1, 14));
    if (sh.getLastColumn() > 0) sh.getRange(1,1,1,sh.getLastColumn())
      .setFontWeight('bold').setBackground('#17324D').setFontColor('#FFFFFF').setWrap(true);
  });
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + Number(days));
  return d;
}
function atHour_(date, hour) {
  const d = new Date(date); d.setHours(hour,0,0,0); return d;
}
