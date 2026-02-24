// ═══════════════════════════════════════════════════════════════════
// WATERFRONT ROTA MANAGER — Google Apps Script Backend
// Deploy as Web App: Execute as ME, Access = Anyone
// ═══════════════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────────────
const SPREADSHEET_ID = '14Ig1OgLPvD3rXQrOeb04p_bTil5nBYeqpweYDzv8m34';
const MANAGER_EMAILS = ['quabsey@gmail.com']; // Add more manager emails as needed
const FROM_NAME = 'Waterfront Rota System';

// ── PERFORMANCE: Spreadsheet cache (per execution) ─────────────────
let _ssCache = null;
function getSpreadsheet() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

// Helper: normalize cell value to string (handles Google Sheets Date auto-conversion)
function cellStr(val) {
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  return String(val);
}

// Sheet names
const SHEET_ROTA = 'Rota';
const SHEET_STAFF = 'Staff';
const SHEET_LEAVE = 'Leave';
const SHEET_THEATRE = 'Theatre';
const SHEET_LOG = 'EmailLog';
const SHEET_CONFIG = 'Config';

// ── WEB APP ENTRY POINTS ──────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter;
  const action = params.action;

  try {
    let result;
    // Use lock for write operations to prevent race conditions
    const writeActions = ['addStaff','updateStaff','deleteStaff','saveShift','saveMultipleShifts','copyWeek','saveTheatre','submitLeave','approveLeave','rejectLeave','publishRota','saveConfig'];
    const needsLock = writeActions.indexOf(action) >= 0;
    let lock = null;

    if (needsLock) {
      lock = LockService.getScriptLock();
      lock.tryLock(10000); // Wait up to 10 seconds
    }

    try {
      switch (action) {
        // Staff
        case 'getStaff': result = getStaff(); break;
        case 'addStaff': result = addStaff(JSON.parse(params.data)); break;
        case 'updateStaff': result = updateStaff(JSON.parse(params.data)); break;
        case 'deleteStaff': result = deleteStaff(params.id); break;

        // Rota
        case 'getRota': result = getRota(params.weekKey); break;
        case 'saveShift': result = saveShift(JSON.parse(params.data)); break;
        case 'saveMultipleShifts': result = saveMultipleShifts(JSON.parse(params.data)); break;
        case 'copyWeek': result = copyWeek(params.fromWeek, params.toWeek); break;

        // Theatre
        case 'getTheatre': result = getTheatre(params.weekKey); break;
        case 'saveTheatre': result = saveTheatre(JSON.parse(params.data)); break;

        // Leave
        case 'getLeave': result = getLeave(); break;
        case 'submitLeave': result = submitLeave(JSON.parse(params.data)); break;
        case 'approveLeave': result = approveLeave(params.id); break;
        case 'rejectLeave': result = rejectLeave(params.id); break;

        // Publish & Notify
        case 'publishRota': result = publishRota(params.weekKey); break;

        // Email Log
        case 'getEmailLog': result = getEmailLog(); break;

        // Config / Settings
        case 'getConfig': result = getAppConfig(); break;
        case 'saveConfig': result = saveAppConfig(JSON.parse(params.config)); break;

        default: result = { error: 'Unknown action: ' + action };
      }
    } finally {
      if (lock) lock.releaseLock();
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error in ' + action + ': ' + err.message + '\n' + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── STAFF FUNCTIONS ───────────────────────────────────────────────
function getStaff() {
  const sheet = getSheet(SHEET_STAFF);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const staff = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // Skip empty rows
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    row._row = i + 1; // Track row number for updates
    staff.push(row);
  }
  return { staff };
}

function addStaff(staffData) {
  const sheet = getSheet(SHEET_STAFF);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = headers.map(h => staffData[h] || '');
  // Generate ID
  newRow[0] = Utilities.getUuid().substring(0, 8);
  sheet.appendRow(newRow);
  return { success: true, id: newRow[0] };
}

function updateStaff(staffData) {
  const sheet = getSheet(SHEET_STAFF);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(staffData.id)) {
      headers.forEach((h, j) => {
        if (staffData[h] !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(staffData[h]);
        }
      });
      return { success: true };
    }
  }
  return { error: 'Staff not found' };
}

function deleteStaff(id) {
  const sheet = getSheet(SHEET_STAFF);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Staff not found' };
}

// ── ROTA FUNCTIONS ────────────────────────────────────────────────
function getRota(weekKey) {
  const sheet = getSheet(SHEET_ROTA);
  const data = sheet.getDataRange().getValues();
  const shifts = {};

  // Format: weekKey | staffId | Mon | Tue | Wed | Thu | Fri | Sat | Sun | published
  for (let i = 1; i < data.length; i++) {
    if (cellStr(data[i][0]) === weekKey) {
      const staffId = String(data[i][1]);
      shifts[staffId] = {
        Mon: data[i][2] || '',
        Tue: data[i][3] || '',
        Wed: data[i][4] || '',
        Thu: data[i][5] || '',
        Fri: data[i][6] || '',
        Sat: data[i][7] || '',
        Sun: data[i][8] || '',
      };
    }
  }

  // Check if published
  const configSheet = getSheet(SHEET_CONFIG);
  const configData = configSheet.getDataRange().getValues();
  let published = false;
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'published_' + weekKey) {
      published = configData[i][1] === true || configData[i][1] === 'TRUE';
      break;
    }
  }

  return { shifts, published, weekKey };
}

function saveShift(data) {
  // data = { weekKey, staffId, day, value }
  const sheet = getSheet(SHEET_ROTA);
  const allData = sheet.getDataRange().getValues();
  const dayIndex = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].indexOf(data.day) + 2; // +2 for weekKey and staffId columns

  // Find existing row
  for (let i = 1; i < allData.length; i++) {
    if (cellStr(allData[i][0]) === data.weekKey && String(allData[i][1]) === String(data.staffId)) {
      sheet.getRange(i + 1, dayIndex + 1).setValue(data.value);

      // If rota is published, send change notification
      checkAndNotifyChange(data);
      return { success: true, updated: true };
    }
  }

  // Create new row
  const newRow = [data.weekKey, data.staffId, '', '', '', '', '', '', ''];
  newRow[dayIndex] = data.value;
  sheet.appendRow(newRow);
  return { success: true, created: true };
}

function saveMultipleShifts(data) {
  // data = { weekKey, shifts: { staffId: { Mon: '...', Tue: '...', ... } } }
  const sheet = getSheet(SHEET_ROTA);
  const allData = sheet.getDataRange().getValues();
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

  Object.entries(data.shifts).forEach(([staffId, shifts]) => {
    let found = false;
    for (let i = 1; i < allData.length; i++) {
      if (cellStr(allData[i][0]) === data.weekKey && String(allData[i][1]) === String(staffId)) {
        days.forEach((d, di) => {
          if (shifts[d] !== undefined) {
            sheet.getRange(i + 1, di + 3).setValue(shifts[d]);
          }
        });
        found = true;
        break;
      }
    }
    if (!found) {
      const newRow = [data.weekKey, staffId];
      days.forEach(d => newRow.push(shifts[d] || ''));
      sheet.appendRow(newRow);
    }
  });

  return { success: true };
}

function copyWeek(fromWeek, toWeek) {
  const fromData = getRota(fromWeek);
  if (!fromData.shifts || Object.keys(fromData.shifts).length === 0) {
    return { error: 'No rota data found for source week' };
  }
  return saveMultipleShifts({ weekKey: toWeek, shifts: fromData.shifts });
}

// ── THEATRE FUNCTIONS ─────────────────────────────────────────────
function getTheatre(weekKey) {
  const sheet = getSheet(SHEET_THEATRE);
  const data = sheet.getDataRange().getValues();
  const schedule = {};

  // Format: weekKey | room | Mon | Tue | Wed | Thu | Fri | Sat | Sun
  for (let i = 1; i < data.length; i++) {
    if (cellStr(data[i][0]) === weekKey) {
      schedule[data[i][1]] = {
        Mon: data[i][2] || '',
        Tue: data[i][3] || '',
        Wed: data[i][4] || '',
        Thu: data[i][5] || '',
        Fri: data[i][6] || '',
        Sat: data[i][7] || '',
        Sun: data[i][8] || '',
      };
    }
  }
  return { schedule, weekKey };
}

function saveTheatre(data) {
  // data = { weekKey, room, day, value }
  const sheet = getSheet(SHEET_THEATRE);
  const allData = sheet.getDataRange().getValues();
  const dayIndex = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].indexOf(data.day) + 2;

  for (let i = 1; i < allData.length; i++) {
    if (cellStr(allData[i][0]) === data.weekKey && allData[i][1] === data.room) {
      sheet.getRange(i + 1, dayIndex + 1).setValue(data.value);
      return { success: true };
    }
  }

  const newRow = [data.weekKey, data.room, '', '', '', '', '', '', ''];
  newRow[dayIndex] = data.value;
  sheet.appendRow(newRow);
  return { success: true };
}

// ── LEAVE FUNCTIONS ───────────────────────────────────────────────
function getLeave() {
  const sheet = getSheet(SHEET_LEAVE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const requests = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    row._row = i + 1;
    requests.push(row);
  }
  return { requests };
}

function submitLeave(data) {
  // data = { staffId, staffName, startDate, endDate, type, reason }
  const sheet = getSheet(SHEET_LEAVE);
  const id = Utilities.getUuid().substring(0, 8);
  const now = new Date().toISOString();

  sheet.appendRow([
    id,
    data.staffId,
    data.staffName,
    data.type || 'Annual Leave',
    data.startDate,
    data.endDate,
    data.reason || '',
    'pending',
    now,
    '' // approvedDate
  ]);

  // Notify managers (use dynamic config, fallback to hardcoded)
  const managerEmails = getManagerEmails();
  managerEmails.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: `Leave Request: ${data.staffName} — ${data.type}`,
        htmlBody: buildLeaveRequestEmail(data),
        name: FROM_NAME
      });
    } catch (e) {
      Logger.log('Failed to send to ' + email + ': ' + e.message);
    }
  });

  logEmail('Leave request notification', managerEmails.join(', '), data.staffName + ' requested ' + data.type);

  return { success: true, id };
}

function approveLeave(id) {
  return updateLeaveStatus(id, 'approved');
}

function rejectLeave(id) {
  return updateLeaveStatus(id, 'rejected');
}

function updateLeaveStatus(id, status) {
  const sheet = getSheet(SHEET_LEAVE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('status') + 1;
  const approvedDateCol = headers.indexOf('approvedDate') + 1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      sheet.getRange(i + 1, approvedDateCol).setValue(new Date().toISOString());

      // Get staff details for notification
      const staffName = data[i][headers.indexOf('staffName')];
      const staffId = data[i][headers.indexOf('staffId')];
      const startDate = data[i][headers.indexOf('startDate')];
      const endDate = data[i][headers.indexOf('endDate')];
      const type = data[i][headers.indexOf('type')];

      // Find staff email
      const staffSheet = getSheet(SHEET_STAFF);
      const staffData = staffSheet.getDataRange().getValues();
      const staffHeaders = staffData[0];
      const emailCol = staffHeaders.indexOf('email');
      const idCol = staffHeaders.indexOf('id');

      for (let j = 1; j < staffData.length; j++) {
        if (String(staffData[j][idCol]) === String(staffId) && staffData[j][emailCol]) {
          try {
            MailApp.sendEmail({
              to: staffData[j][emailCol],
              subject: `Leave ${status === 'approved' ? 'Approved' : 'Rejected'}: ${startDate} to ${endDate}`,
              htmlBody: buildLeaveDecisionEmail(staffName, type, startDate, endDate, status),
              name: FROM_NAME
            });
            logEmail('Leave ' + status, staffData[j][emailCol], type + ': ' + startDate + ' to ' + endDate);
          } catch (e) {
            Logger.log('Failed to send leave notification: ' + e.message);
          }
          break;
        }
      }

      return { success: true };
    }
  }
  return { error: 'Leave request not found' };
}

// ── PUBLISH & NOTIFY ──────────────────────────────────────────────
function publishRota(weekKey) {
  // Mark as published
  const configSheet = getSheet(SHEET_CONFIG);
  const configData = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'published_' + weekKey) {
      configSheet.getRange(i + 1, 2).setValue(true);
      configSheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      found = true;
      break;
    }
  }
  if (!found) {
    configSheet.appendRow(['published_' + weekKey, true, new Date().toISOString()]);
  }

  // Get rota data
  const rotaData = getRota(weekKey);
  const staffResult = getStaff();
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  let emailsSent = 0;

  // Send individual emails to each staff member with shifts
  staffResult.staff.forEach(staff => {
    const shifts = rotaData.shifts[staff.id];
    if (!shifts || !staff.email) return;

    const hasShifts = days.some(d => shifts[d] && shifts[d] !== '');
    if (!hasShifts) return;

    try {
      MailApp.sendEmail({
        to: staff.email,
        subject: `Your Rota — Week of ${weekKey}`,
        htmlBody: buildRotaEmail(staff, shifts, weekKey),
        name: FROM_NAME
      });
      emailsSent++;
      logEmail('Rota published', staff.email, 'Week of ' + weekKey);
    } catch (e) {
      Logger.log('Failed to send rota to ' + staff.email + ': ' + e.message);
    }
  });

  return { success: true, emailsSent, weekKey };
}

function checkAndNotifyChange(data) {
  // Check if this week's rota is published
  const configSheet = getSheet(SHEET_CONFIG);
  const configData = configSheet.getDataRange().getValues();
  let isPublished = false;
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'published_' + data.weekKey) {
      isPublished = configData[i][1] === true || configData[i][1] === 'TRUE';
      break;
    }
  }

  if (!isPublished) return;

  // Find staff email
  const staffResult = getStaff();
  const staff = staffResult.staff.find(s => String(s.id) === String(data.staffId));
  if (!staff || !staff.email) return;

  try {
    MailApp.sendEmail({
      to: staff.email,
      subject: `Shift Change — ${data.day}, week of ${data.weekKey}`,
      htmlBody: buildShiftChangeEmail(staff.name, data.day, data.value, data.weekKey),
      name: FROM_NAME
    });
    logEmail('Shift change', staff.email, data.day + ': ' + data.value);
  } catch (e) {
    Logger.log('Failed to send shift change notification: ' + e.message);
  }
}

// ── EMAIL TEMPLATES ───────────────────────────────────────────────
function buildRotaEmail(staff, shifts, weekKey) {
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const fullDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  let shiftRows = '';
  days.forEach((d, i) => {
    const val = shifts[d] || '—';
    const bg = val === '—' || !val ? '#F9FAFB' : val.toUpperCase().startsWith('DO') ? '#F3F4F6' : val.toUpperCase().startsWith('A/L') ? '#FEF3C7' : '#ECFDF5';
    shiftRows += `<tr><td style="padding:10px 14px;border-bottom:1px solid #E5E7EB;font-weight:600;color:#374151">${fullDays[i]}</td><td style="padding:10px 14px;border-bottom:1px solid #E5E7EB;background:${bg};text-align:center;font-weight:600;color:#111827">${val}</td></tr>`;
  });

  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">
      <div style="background:linear-gradient(135deg,#1E3A5F,#1E40AF);padding:20px 24px;color:#fff">
        <h2 style="margin:0;font-size:18px">Your Rota</h2>
        <p style="margin:4px 0 0;opacity:0.8;font-size:14px">Week of ${weekKey}</p>
      </div>
      <div style="padding:20px 24px">
        <p style="color:#374151;font-size:14px;margin:0 0 16px">Hi ${staff.name},</p>
        <p style="color:#374151;font-size:14px;margin:0 0 16px">Here are your shifts for the upcoming week:</p>
        <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;border:1px solid #E5E7EB">
          <thead><tr style="background:#1E3A5F"><th style="padding:10px 14px;color:#fff;text-align:left;font-size:13px">Day</th><th style="padding:10px 14px;color:#fff;text-align:center;font-size:13px">Shift</th></tr></thead>
          <tbody>${shiftRows}</tbody>
        </table>
        <p style="color:#6B7280;font-size:13px;margin:16px 0 0">If you have any questions about your rota, please contact your manager.</p>
      </div>
      <div style="background:#F9FAFB;padding:14px 24px;border-top:1px solid #E5E7EB">
        <p style="color:#9CA3AF;font-size:12px;margin:0;text-align:center">Waterfront Private Hospital — Rota Management System</p>
      </div>
    </div>`;
}

function buildShiftChangeEmail(name, day, newShift, weekKey) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">
      <div style="background:#F59E0B;padding:20px 24px;color:#fff">
        <h2 style="margin:0;font-size:18px">Shift Change Notice</h2>
        <p style="margin:4px 0 0;opacity:0.9;font-size:14px">Week of ${weekKey}</p>
      </div>
      <div style="padding:20px 24px">
        <p style="color:#374151;font-size:14px;margin:0 0 12px">Hi ${name},</p>
        <p style="color:#374151;font-size:14px;margin:0 0 16px">Your shift on <strong>${day}</strong> has been updated:</p>
        <div style="background:#FEF9C3;border:1px solid #FCD34D;border-radius:8px;padding:14px 18px;text-align:center">
          <span style="font-size:18px;font-weight:700;color:#92400E">${newShift || 'Shift removed'}</span>
        </div>
        <p style="color:#6B7280;font-size:13px;margin:16px 0 0">Please contact your manager if you have any concerns.</p>
      </div>
    </div>`;
}

function buildLeaveRequestEmail(data) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">
      <div style="background:#8B5CF6;padding:20px 24px;color:#fff">
        <h2 style="margin:0;font-size:18px">Leave Request</h2>
      </div>
      <div style="padding:20px 24px">
        <p style="margin:0 0 12px"><strong>${data.staffName}</strong> has requested:</p>
        <table style="width:100%">
          <tr><td style="padding:6px 0;color:#6B7280">Type:</td><td style="padding:6px 0;font-weight:600">${data.type}</td></tr>
          <tr><td style="padding:6px 0;color:#6B7280">From:</td><td style="padding:6px 0;font-weight:600">${data.startDate}</td></tr>
          <tr><td style="padding:6px 0;color:#6B7280">To:</td><td style="padding:6px 0;font-weight:600">${data.endDate}</td></tr>
          ${data.reason ? '<tr><td style="padding:6px 0;color:#6B7280">Reason:</td><td style="padding:6px 0">' + data.reason + '</td></tr>' : ''}
        </table>
        <p style="color:#6B7280;font-size:13px;margin:16px 0 0">Log in to the rota system to approve or reject this request.</p>
      </div>
    </div>`;
}

function buildLeaveDecisionEmail(name, type, startDate, endDate, status) {
  const color = status === 'approved' ? '#10B981' : '#EF4444';
  const label = status === 'approved' ? 'Approved' : 'Rejected';
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">
      <div style="background:${color};padding:20px 24px;color:#fff">
        <h2 style="margin:0;font-size:18px">Leave ${label}</h2>
      </div>
      <div style="padding:20px 24px">
        <p style="color:#374151;font-size:14px;margin:0 0 12px">Hi ${name},</p>
        <p style="color:#374151;font-size:14px;margin:0 0 16px">Your ${type} request has been <strong>${status}</strong>:</p>
        <div style="background:#F3F4F6;border-radius:8px;padding:14px 18px">
          <p style="margin:0;font-size:14px"><strong>${startDate}</strong> to <strong>${endDate}</strong></p>
        </div>
      </div>
    </div>`;
}

// ── CONFIG / SETTINGS ─────────────────────────────────────────────
function getAppConfig() {
  const sheet = getSheet(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = { manager_emails: MANAGER_EMAILS, auto_email: true };

  for (let i = 1; i < data.length; i++) {
    const key = data[i][0], value = data[i][1];
    if (key === 'manager_emails' && value) {
      try { config.manager_emails = JSON.parse(value); } catch (e) {}
    }
    if (key === 'auto_email') {
      config.auto_email = value === true || value === 'TRUE' || value === 'true';
    }
  }
  return { success: true, config };
}

function saveAppConfig(config) {
  const sheet = getSheet(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  let foundManagers = false, foundAutoEmail = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'manager_emails') {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(config.manager_emails || []));
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      foundManagers = true;
    }
    if (data[i][0] === 'auto_email') {
      sheet.getRange(i + 1, 2).setValue(config.auto_email);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      foundAutoEmail = true;
    }
  }
  if (!foundManagers) sheet.appendRow(['manager_emails', JSON.stringify(config.manager_emails || []), new Date().toISOString()]);
  if (!foundAutoEmail) sheet.appendRow(['auto_email', config.auto_email, new Date().toISOString()]);

  return { success: true };
}

function getManagerEmails() {
  const result = getAppConfig();
  return result.config.manager_emails && result.config.manager_emails.length > 0
    ? result.config.manager_emails
    : MANAGER_EMAILS;
}

// ── EMAIL LOG ─────────────────────────────────────────────────────
function logEmail(type, recipient, details) {
  try {
    const sheet = getSheet(SHEET_LOG);
    sheet.appendRow([new Date().toISOString(), type, recipient, details]);
  } catch (e) {
    Logger.log('Failed to log email: ' + e.message);
  }
}

function getEmailLog() {
  const sheet = getSheet(SHEET_LOG);
  const data = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    logs.push({ timestamp: data[i][0], type: data[i][1], recipient: data[i][2], details: data[i][3] });
  }
  // Return newest first
  logs.reverse();
  return { logs };
}

// ── WEEKLY REMINDER (Time-driven trigger) ─────────────────────────
// To enable: Run setupWeeklyReminder() once. It creates a Sunday 6pm trigger.
function setupWeeklyReminder() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyReminders') ScriptApp.deleteTrigger(t);
  });
  // Create new Sunday 6pm trigger
  ScriptApp.newTrigger('sendWeeklyReminders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(18)
    .create();
  Logger.log('Weekly reminder trigger created for Sundays at 6pm');
}

function sendWeeklyReminders() {
  // Calculate next Monday's date
  const today = new Date();
  const daysUntilMonday = (8 - today.getDay()) % 7 || 7;
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilMonday);
  const weekKey = nextMonday.toISOString().slice(0, 10);

  const rotaData = getRota(weekKey);
  const staffResult = getStaff();
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  let sent = 0;

  staffResult.staff.forEach(staff => {
    const shifts = rotaData.shifts[staff.id];
    if (!shifts || !staff.email) return;
    const hasShifts = days.some(d => shifts[d] && shifts[d] !== '');
    if (!hasShifts) return;

    try {
      MailApp.sendEmail({
        to: staff.email,
        subject: `Reminder: Your Rota — Week of ${weekKey}`,
        htmlBody: buildRotaEmail(staff, shifts, weekKey),
        name: FROM_NAME
      });
      sent++;
    } catch (e) {
      Logger.log('Reminder failed for ' + staff.email + ': ' + e.message);
    }
  });

  logEmail('Weekly reminder', 'All staff', sent + ' reminders sent for week of ' + weekKey);
  Logger.log('Sent ' + sent + ' weekly reminders for ' + weekKey);
}

// ── SETUP: Run once to create sheet structure ─────────────────────
function setupSheets() {
  const ss = getSpreadsheet();

  // Staff sheet
  let sheet = ss.getSheetByName(SHEET_STAFF);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_STAFF);
    sheet.appendRow(['id', 'name', 'group', 'department', 'hours', 'contractHours', 'email', 'phone']);

    // Pre-populate with your current staff
    const staff = [
      // Admin
      ['adm-01', 'Janey', 'admin', '', 'FULL TIME', 37.5, '', ''],
      ['adm-02', 'Katrina', 'admin', '', 'FULL TIME', 37.5, '', ''],
      ['adm-03', 'Debbie', 'admin', '', 'PART TIME', 22.5, '', ''],
      ['adm-04', 'Danielle', 'admin', '', 'PART TIME', 15.5, '', ''],
      ['adm-05', 'Margaret', 'admin', '', 'PART TIME', 15, '', ''],
      ['adm-06', 'Lucy', 'admin', '', 'BANK', 0, '', ''],
      // Clinical
      ['cli-01', 'Morag', 'clinical', '', '', 37.5, '', ''],
      ['cli-02', 'Islay', 'clinical', '', '', 25, '', ''],
      ['cli-03', 'Oksana', 'clinical', '', '', 37.5, '', ''],
      ['cli-04', 'Christina', 'clinical', '', '', 37.5, '', ''],
      ['cli-05', 'Cezary', 'clinical', '', '', 37.5, '', ''],
      ['cli-06', 'Piel', 'clinical', '', '', 37.5, '', ''],
      ['cli-07', 'Inez', 'clinical', '', '', 30, '', ''],
      ['cli-08', 'Laura', 'clinical', '', '', 36, '', ''],
      ['cli-09', 'Kerry Ann', 'clinical', '', '', 33, '', ''],
      ['cli-10', 'Wojtek', 'clinical', '', '', 37.5, '', ''],
      // Bank
      ['bnk-01', 'Gleb', 'bank', '', '', 0, '', ''],
      ['bnk-02', 'Rona', 'bank', '', '', 0, '', ''],
      ['bnk-03', 'Sybil', 'bank', '', '', 0, '', ''],
      ['bnk-04', 'Gomes', 'bank', '', '', 0, '', ''],
      ['bnk-05', 'Myron', 'bank', '', '', 0, '', ''],
      ['bnk-06', 'Claudia', 'bank', 'OPD', '', 0, '', ''],
      ['bnk-07', 'Louise', 'bank', 'OPD', '', 0, '', ''],
      ['bnk-08', 'Sally', 'bank', 'OPD/Ward', '', 0, '', ''],
      ['bnk-09', 'Jo', 'bank', 'OPD', '', 0, '', ''],
      ['bnk-10', 'Eve', 'bank', 'Ward', '', 0, '', ''],
      ['bnk-11', 'Damien', 'bank', 'Ward', '', 0, '', ''],
      ['bnk-12', 'Draga', 'bank', 'Ward', '', 0, '', ''],
      ['bnk-13', 'Irina', 'bank', 'Ward', '', 0, '', ''],
      ['bnk-14', 'Olga', 'bank', 'ODP', '', 0, '', ''],
      ['bnk-15', 'Ashleigh', 'bank', 'ODP', '', 0, '', ''],
      ['bnk-16', 'Steve G', 'bank', 'ODP', '', 0, '', ''],
      ['bnk-17', 'Elisabeth', 'bank', 'TH', '', 0, '', ''],
      ['bnk-18', 'Clare', 'bank', 'Scrub/Rec', '', 0, '', ''],
      ['bnk-19', 'Isabella', 'bank', 'REC', '', 0, '', ''],
      ['bnk-20', 'Maxine', 'bank', 'ODP', '', 0, '', ''],
      ['bnk-21', 'Fiona B', 'bank', 'Scrub', '', 0, '', ''],
      ['bnk-22', 'Carl', 'bank', 'TH', '', 0, '', ''],
      ['bnk-23', 'Mariam', 'bank', '', '', 0, '', ''],
    ];
    staff.forEach(row => sheet.appendRow(row));
  }

  // Rota sheet
  if (!ss.getSheetByName(SHEET_ROTA)) {
    const rotaSheet = ss.insertSheet(SHEET_ROTA);
    rotaSheet.appendRow(['weekKey', 'staffId', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']);
  }

  // Theatre sheet
  if (!ss.getSheetByName(SHEET_THEATRE)) {
    const theatreSheet = ss.insertSheet(SHEET_THEATRE);
    theatreSheet.appendRow(['weekKey', 'room', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']);
  }

  // Leave sheet
  if (!ss.getSheetByName(SHEET_LEAVE)) {
    const leaveSheet = ss.insertSheet(SHEET_LEAVE);
    leaveSheet.appendRow(['id', 'staffId', 'staffName', 'type', 'startDate', 'endDate', 'reason', 'status', 'submittedDate', 'approvedDate']);
  }

  // Email log sheet
  if (!ss.getSheetByName(SHEET_LOG)) {
    const logSheet = ss.insertSheet(SHEET_LOG);
    logSheet.appendRow(['timestamp', 'type', 'recipient', 'details']);
  }

  // Config sheet
  if (!ss.getSheetByName(SHEET_CONFIG)) {
    const configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(['key', 'value', 'timestamp']);
  }

  Logger.log('All sheets created successfully!');
  return { success: true, message: 'Sheets created and populated with staff data' };
}
