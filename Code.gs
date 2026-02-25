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

// ── PIN VALIDATION ────────────────────────────────────────────────
function validatePin(pin) {
  const sheet = getSheet(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'admin_pin') {
      return String(data[i][1]) === String(pin);
    }
  }
  return true; // No PIN configured — allow all writes (backward compatible)
}

function isPinConfigured() {
  const sheet = getSheet(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'admin_pin' && String(data[i][1]).trim() !== '') return true;
  }
  return false;
}

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
    const writeActions = ['addStaff','updateStaff','deleteStaff','saveShift','saveMultipleShifts','copyWeek','saveTheatre','submitLeave','approveLeave','rejectLeave','publishRota','publishMonthlyRota','publishMonthlyOverview','saveConfig'];
    const needsLock = writeActions.indexOf(action) >= 0;
    let lock = null;

    // Validate PIN for write operations
    if (needsLock) {
      if (!validatePin(params.pin || '')) {
        return ContentService
          .createTextOutput(JSON.stringify({ error: 'INVALID_PIN', message: 'Invalid admin PIN' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      lock = LockService.getScriptLock();
      lock.tryLock(10000);
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
        case 'getLeaveBalances': result = getLeaveBalances(); break;
        case 'submitLeave': result = submitLeave(JSON.parse(params.data)); break;
        case 'approveLeave': result = approveLeave(params.id); break;
        case 'rejectLeave': result = rejectLeave(params.id); break;

        // Publish & Notify
        case 'publishRota': result = publishRota(params.weekKey, params.emailMode || 'all'); break;
        case 'publishMonthlyRota': result = publishMonthlyRota(params.staffId, params.startWeekKey); break;
        case 'publishMonthlyOverview': result = publishMonthlyOverview(params.startWeekKey); break;

        // Email Log
        case 'getEmailLog': result = getEmailLog(); break;

        // Config / Settings
        case 'getConfig': result = getAppConfig(); break;
        case 'saveConfig': result = saveAppConfig(JSON.parse(params.config)); break;

        // PIN verification (no lock needed but listed for frontend use)
        case 'verifyPin': result = { valid: validatePin(params.pin || '') }; break;

        // Combined initial load
        case 'getInit': result = getInit(params.weekKey); break;

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

// ── COMBINED INITIAL LOAD (single request for all startup data) ───
function getInit(weekKey) {
  // Fetch everything the app needs to render in ONE call
  // This eliminates multiple cold-start round-trips
  const staffResult = getStaff();
  const leaveResult = getLeave();
  const rotaResult = getRota(weekKey);
  const theatreResult = getTheatre(weekKey);
  const configResult = getAppConfig();
  const balancesResult = getLeaveBalances();
  return {
    staff: staffResult.staff || [],
    leave: leaveResult.requests || [],
    rota: rotaResult.shifts || {},
    published: rotaResult.published || false,
    theatre: theatreResult.schedule || {},
    config: configResult.config || {},
    has_pin: isPinConfigured(),
    leaveBalances: balancesResult.balances || [],
    timestamp: new Date().toISOString()
  };
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
    headers.forEach((h, j) => { row[String(h).toLowerCase().trim()] = data[i][j]; });
    row._row = i + 1; // Track row number for updates
    staff.push(row);
  }
  return { staff };
}

function addStaff(staffData) {
  const sheet = getSheet(SHEET_STAFF);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = headers.map(h => staffData[String(h).toLowerCase().trim()] || '');
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
        const key = String(h).toLowerCase().trim();
        if (staffData[key] !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(staffData[key]);
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

// ── LEAVE HELPERS ─────────────────────────────────────────────────
function getLeaveShiftCode(leaveType) {
  const map = {
    'Annual Leave': 'A/L',
    'Study Leave': 'S/L',
    'Sick Leave': 'SICK',
    'Compassionate Leave': 'C/L'
  };
  return map[leaveType] || 'A/L';
}

function getWeekdaysBetween(startDateStr, endDateStr) {
  const result = [];
  const DAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const start = new Date(startDateStr + 'T00:00:00');
  const end = new Date(endDateStr + 'T00:00:00');
  const current = new Date(start);
  while (current <= end) {
    const dow = current.getDay();
    if (dow >= 1 && dow <= 5) { // Mon-Fri
      // Calculate the Monday (weekKey) for this date
      const monday = new Date(current);
      monday.setDate(current.getDate() - (dow - 1));
      result.push({
        date: current.toISOString().slice(0, 10),
        dayName: DAYS[dow],
        weekKey: monday.toISOString().slice(0, 10)
      });
    }
    current.setDate(current.getDate() + 1);
  }
  return result;
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

      // Auto-insert shifts into rota when approved
      if (status === 'approved') {
        const shiftCode = getLeaveShiftCode(type);
        const weekdays = getWeekdaysBetween(cellStr(startDate), cellStr(endDate));
        weekdays.forEach(function(wd) {
          saveShift({ weekKey: wd.weekKey, staffId: String(staffId), day: wd.dayName, value: shiftCode });
        });
      }

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
              subject: `Leave ${status === 'approved' ? 'Approved' : 'Rejected'}: ${cellStr(startDate)} to ${cellStr(endDate)}`,
              htmlBody: buildLeaveDecisionEmail(staffName, type, cellStr(startDate), cellStr(endDate), status),
              name: FROM_NAME
            });
            logEmail('Leave ' + status, staffData[j][emailCol], type + ': ' + cellStr(startDate) + ' to ' + cellStr(endDate));
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

// ── LEAVE BALANCES ────────────────────────────────────────────────
function getLeaveBalances() {
  const staffResult = getStaff();
  const leaveResult = getLeave();
  const allStaff = staffResult.staff || [];
  const allLeave = leaveResult.requests || [];
  const today = new Date();
  const balances = [];

  // Only admin + clinical staff
  const eligibleStaff = allStaff.filter(function(s) {
    return s.group === 'admin' || s.group === 'clinical';
  });

  eligibleStaff.forEach(function(staff) {
    var entitlement = parseInt(staff.leaveentitlement) || 28;
    var startDateStr = staff.startdate ? cellStr(staff.startdate) : '';

    // Calculate leave year based on start date anniversary
    var yearStart, yearEnd;
    if (startDateStr) {
      var sd = new Date(startDateStr + 'T00:00:00');
      yearStart = new Date(today.getFullYear(), sd.getMonth(), sd.getDate());
      if (yearStart > today) {
        yearStart.setFullYear(yearStart.getFullYear() - 1);
      }
      yearEnd = new Date(yearStart);
      yearEnd.setFullYear(yearEnd.getFullYear() + 1);
      yearEnd.setDate(yearEnd.getDate() - 1);
    } else {
      // Default: calendar year
      yearStart = new Date(today.getFullYear(), 0, 1);
      yearEnd = new Date(today.getFullYear(), 11, 31);
    }

    var yearStartStr = yearStart.toISOString().slice(0, 10);
    var yearEndStr = yearEnd.toISOString().slice(0, 10);

    // Count approved leave days in this leave year
    var annualTaken = 0;
    var studyDays = 0;
    var sickDays = 0;
    var compassionateDays = 0;

    allLeave.forEach(function(req) {
      if (String(req.staffId) !== String(staff.id)) return;
      if (req.status !== 'approved') return;

      var leaveStart = cellStr(req.startDate);
      var leaveEnd = cellStr(req.endDate);

      // Check if leave overlaps with leave year
      if (leaveEnd < yearStartStr || leaveStart > yearEndStr) return;

      // Clamp to leave year boundaries
      var effectiveStart = leaveStart < yearStartStr ? yearStartStr : leaveStart;
      var effectiveEnd = leaveEnd > yearEndStr ? yearEndStr : leaveEnd;

      var weekdays = getWeekdaysBetween(effectiveStart, effectiveEnd);
      var dayCount = weekdays.length;

      var leaveType = req.type || 'Annual Leave';
      if (leaveType === 'Annual Leave') annualTaken += dayCount;
      else if (leaveType === 'Study Leave') studyDays += dayCount;
      else if (leaveType === 'Sick Leave') sickDays += dayCount;
      else if (leaveType === 'Compassionate Leave') compassionateDays += dayCount;
    });

    balances.push({
      staffId: staff.id,
      staffName: staff.name,
      group: staff.group,
      entitlement: entitlement,
      taken: annualTaken,
      remaining: entitlement - annualTaken,
      yearStart: yearStartStr,
      yearEnd: yearEndStr,
      leaveBreakdown: {
        study: studyDays,
        sick: sickDays,
        compassionate: compassionateDays
      }
    });
  });

  return { balances: balances };
}

// ── PUBLISH & NOTIFY ──────────────────────────────────────────────
function publishRota(weekKey, emailMode) {
  emailMode = emailMode || 'all'; // 'all' | 'staff-only' | 'none'

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

  // If no emails requested, return early
  if (emailMode === 'none') {
    return { success: true, emailsSent: 0, managerEmailsSent: 0, weekKey, debug: ['Marked as published, no emails sent'], errors: [] };
  }

  const rotaData = getRota(weekKey);
  const staffResult = getStaff();
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  let emailsSent = 0;
  let managerEmailsSent = 0;
  const errors = [];
  const debug = [];
  const emailQuota = MailApp.getRemainingDailyQuota();
  debug.push(staffResult.staff.length + ' staff, ' + emailQuota + ' emails remaining in quota');

  // Send individual emails to each staff member with shifts
  staffResult.staff.forEach(staff => {
    const shifts = rotaData.shifts[staff.id];
    if (!shifts || !staff.email || String(staff.email).trim() === '') return;
    const hasShifts = days.some(d => shifts[d] && shifts[d] !== '');
    if (!hasShifts) return;

    try {
      MailApp.sendEmail({
        to: staff.email,
        subject: 'Your Rota — Week of ' + weekKey,
        htmlBody: buildRotaEmail(staff, shifts, weekKey),
        name: FROM_NAME
      });
      emailsSent++;
      logEmail('Rota published', staff.email, 'Week of ' + weekKey);
    } catch (e) {
      errors.push(staff.name + ': ' + e.message);
    }
  });

  // Send manager full-rota summary if mode is 'all'
  if (emailMode === 'all') {
    const managerEmails = getManagerEmails();
    if (managerEmails.length > 0) {
      const fullRotaHtml = buildFullRotaEmail(staffResult.staff, rotaData.shifts, weekKey);
      managerEmails.forEach(mEmail => {
        try {
          MailApp.sendEmail({
            to: mEmail,
            subject: 'Full Weekly Rota — Week of ' + weekKey,
            htmlBody: fullRotaHtml,
            name: FROM_NAME
          });
          managerEmailsSent++;
          logEmail('Full rota (manager)', mEmail, 'Week of ' + weekKey);
        } catch (e) {
          errors.push('Manager ' + mEmail + ': ' + e.message);
        }
      });
    }
  }

  debug.push(emailsSent + ' staff emails, ' + managerEmailsSent + ' manager emails');
  return { success: true, emailsSent, managerEmailsSent, weekKey, debug, errors };
}

// ── MONTHLY ROTA EMAIL ────────────────────────────────────────────
function publishMonthlyRota(staffId, startWeekKey) {
  const staffResult = getStaff();
  const staff = staffResult.staff.find(s => String(s.id) === String(staffId));
  if (!staff) return { error: 'Staff not found' };
  if (!staff.email) return { error: staff.name + ' has no email address' };

  // Calculate 5 consecutive Monday dates
  const startDate = new Date(startWeekKey + 'T00:00:00');
  const weekKeys = [];
  for (let w = 0; w < 5; w++) {
    const monday = new Date(startDate);
    monday.setDate(startDate.getDate() + w * 7);
    weekKeys.push(monday.toISOString().slice(0, 10));
  }

  // Fetch rota data for all 5 weeks
  const allWeekShifts = {};
  weekKeys.forEach(wk => {
    const rotaData = getRota(wk);
    allWeekShifts[wk] = rotaData.shifts[staffId] || {};
  });

  const htmlBody = buildMonthlyRotaEmail(staff, allWeekShifts, weekKeys);

  try {
    MailApp.sendEmail({
      to: staff.email,
      subject: 'Your Monthly Rota — Starting ' + startWeekKey,
      htmlBody: htmlBody,
      name: FROM_NAME
    });
    logEmail('Monthly rota', staff.email, '5 weeks from ' + startWeekKey);
    return { success: true, emailsSent: 1, staffName: staff.name };
  } catch (e) {
    return { error: 'Failed to send: ' + e.message };
  }
}

function publishMonthlyOverview(startWeekKey) {
  const managerEmails = getManagerEmails();
  if (!managerEmails || managerEmails.length === 0) return { error: 'No managers configured' };

  const staffResult = getStaff();
  const staffList = staffResult.staff || [];

  // Calculate 5 consecutive Monday dates
  const startDate = new Date(startWeekKey + 'T00:00:00');
  const weekKeys = [];
  for (let w = 0; w < 5; w++) {
    const monday = new Date(startDate);
    monday.setDate(startDate.getDate() + w * 7);
    weekKeys.push(monday.toISOString().slice(0, 10));
  }

  // Fetch rota data for all 5 weeks
  const allWeeksData = {};
  weekKeys.forEach(wk => {
    const rotaData = getRota(wk);
    allWeeksData[wk] = rotaData.shifts || {};
  });

  const htmlBody = buildMonthlyOverviewEmail(staffList, allWeeksData, weekKeys);
  let count = 0;

  managerEmails.forEach(mEmail => {
    try {
      MailApp.sendEmail({
        to: mEmail,
        subject: 'Monthly Rota Overview — 5 Weeks from ' + startWeekKey,
        htmlBody: htmlBody,
        name: FROM_NAME
      });
      logEmail('Monthly overview', mEmail, '5 weeks from ' + startWeekKey);
      count++;
    } catch (e) {
      Logger.log('Failed to send monthly overview to ' + mEmail + ': ' + e.message);
    }
  });

  return { success: true, count: count };
}

function buildMonthlyOverviewEmail(staffList, allWeeksData, weekKeys) {
  const DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  let html = `
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;max-width:900px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">
    <div style="background:linear-gradient(135deg,#1E3A5F,#1E40AF);padding:20px 24px;color:#fff">
      <h1 style="margin:0;font-size:20px;font-weight:700">Monthly Rota Overview</h1>
      <p style="margin:4px 0 0;opacity:0.7;font-size:13px">5-week overview for all staff</p>
    </div>
    <div style="padding:20px 24px">`;

  weekKeys.forEach((wk, wIdx) => {
    const monday = new Date(wk + 'T00:00:00');
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    const weekLabel = monday.getDate() + ' ' + m[monday.getMonth()] + ' – ' + sunday.getDate() + ' ' + m[sunday.getMonth()] + ' ' + sunday.getFullYear();
    const shifts = allWeeksData[wk] || {};

    html += `<h3 style="font-size:14px;font-weight:600;color:#1E40AF;margin:${wIdx > 0 ? '24px' : '0'} 0 8px;padding-bottom:6px;border-bottom:2px solid #1E40AF">${weekLabel}</h3>`;
    html += `<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:8px">
      <thead><tr style="background:#1E3A5F;color:#fff">
        <th style="padding:6px 8px;text-align:left;font-size:10px;font-weight:600">Staff</th>`;
    DAYS.forEach((d, di) => {
      const dt = new Date(monday);
      dt.setDate(monday.getDate() + di);
      html += `<th style="padding:6px 4px;text-align:center;font-size:10px;font-weight:600">${d} ${dt.getDate()}</th>`;
    });
    html += `</tr></thead><tbody>`;

    const groups = [
      { key: 'admin', label: 'ADMIN', color: '#1E40AF' },
      { key: 'clinical', label: 'CLINICAL', color: '#065F46' },
      { key: 'bank', label: 'BANK', color: '#9D174D' }
    ];

    groups.forEach(g => {
      const members = staffList.filter(s => s.group === g.key);
      if (members.length === 0) return;
      html += `<tr><td colspan="8" style="padding:5px 8px;font-weight:700;font-size:10px;color:${g.color};background:#F9FAFB;border-bottom:1px solid ${g.color}">${g.label}</td></tr>`;
      members.forEach(s => {
        const sShifts = shifts[s.id] || {};
        html += `<tr style="border-bottom:1px solid #F3F4F6"><td style="padding:4px 8px;font-weight:500;font-size:11px">${s.name}</td>`;
        DAYS.forEach(d => {
          const val = sShifts[d] || '';
          const bg = val && val.toUpperCase() !== 'DO' ? '#E0F2FE' : (val.toUpperCase() === 'DO' ? '#F3F4F6' : '#fff');
          html += `<td style="padding:4px 3px;text-align:center;font-size:10px;background:${bg}">${val || '—'}</td>`;
        });
        html += '</tr>';
      });
    });

    html += '</tbody></table>';
  });

  html += `
      <p style="font-size:11px;color:#9CA3AF;margin-top:16px;text-align:center">Waterfront Private Hospital Rota Manager</p>
    </div>
    </div>`;

  return html;
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
      subject: 'Shift Change — ' + data.day + ', week of ' + data.weekKey,
      htmlBody: buildShiftChangeEmail(staff.name, data.day, data.value, data.weekKey),
      name: FROM_NAME
    });
    logEmail('Shift change', staff.email, data.day + ': ' + data.value);
  } catch (e) {
    Logger.log('Failed to send shift change notification: ' + e.message);
  }

  // Also notify managers
  const managerEmails = getManagerEmails();
  managerEmails.forEach(mEmail => {
    try {
      MailApp.sendEmail({
        to: mEmail,
        subject: 'Shift Change Alert — ' + staff.name + ' — ' + data.day + ', week of ' + data.weekKey,
        htmlBody: buildShiftChangeEmail(staff.name, data.day, data.value, data.weekKey),
        name: FROM_NAME
      });
    } catch (e) {
      Logger.log('Failed to notify manager of shift change: ' + e.message);
    }
  });
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

// ── FULL ROTA EMAIL (all staff, one week — for managers) ──────────
function buildFullRotaEmail(staffList, allShifts, weekKey) {
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const fullDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const groups = [
    { key: 'admin', label: 'ADMIN' },
    { key: 'clinical', label: 'CLINICAL' },
    { key: 'bank', label: 'BANK' }
  ];

  function shiftBg(val) {
    if (!val || val === '—') return '#F9FAFB';
    var v = String(val).toUpperCase();
    if (v.startsWith('DO')) return '#F3F4F6';
    if (v.startsWith('A/L')) return '#FEF3C7';
    if (v === 'AV') return '#DBEAFE';
    if (v === 'WFH') return '#E0E7FF';
    return '#ECFDF5';
  }

  var dayHeaders = days.map(function(d) {
    return '<th style="padding:6px 4px;background:#1E3A5F;color:#fff;font-size:10px;text-align:center;border:1px solid #E5E7EB">' + d + '</th>';
  }).join('');

  var bodyRows = '';
  groups.forEach(function(g) {
    var members = staffList.filter(function(s) { return s.group === g.key; });
    if (members.length === 0) return;
    bodyRows += '<tr><td colspan="8" style="padding:8px 6px;font-weight:700;font-size:11px;color:#1E40AF;background:#EFF6FF;border:1px solid #E5E7EB">' + g.label + ' (' + members.length + ')</td></tr>';
    members.forEach(function(s) {
      var shifts = allShifts[s.id] || {};
      var cells = days.map(function(d) {
        var val = shifts[d] || '—';
        return '<td style="padding:4px 2px;text-align:center;font-size:9px;font-weight:600;background:' + shiftBg(val) + ';border:1px solid #E5E7EB;color:#111827">' + val + '</td>';
      }).join('');
      bodyRows += '<tr><td style="padding:4px 6px;font-size:10px;font-weight:600;color:#374151;border:1px solid #E5E7EB;white-space:nowrap">' + s.name + '</td>' + cells + '</tr>';
    });
  });

  return '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">' +
    '<div style="background:linear-gradient(135deg,#1E3A5F,#1E40AF);padding:20px 24px;color:#fff">' +
      '<h2 style="margin:0;font-size:18px">Full Weekly Rota</h2>' +
      '<p style="margin:4px 0 0;opacity:0.8;font-size:14px">Week of ' + weekKey + '</p>' +
    '</div>' +
    '<div style="padding:16px">' +
      '<table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif">' +
        '<thead><tr><th style="padding:6px 4px;background:#1E3A5F;color:#fff;font-size:10px;text-align:left;border:1px solid #E5E7EB">Staff</th>' + dayHeaders + '</tr></thead>' +
        '<tbody>' + bodyRows + '</tbody>' +
      '</table>' +
    '</div>' +
    '<div style="background:#F9FAFB;padding:14px 24px;border-top:1px solid #E5E7EB">' +
      '<p style="color:#9CA3AF;font-size:12px;margin:0;text-align:center">Waterfront Private Hospital — Rota Management System</p>' +
    '</div>' +
  '</div>';
}

// ── MONTHLY ROTA EMAIL (one staff, 5 weeks) ───────────────────────
function buildMonthlyRotaEmail(staff, allWeekShifts, weekKeys) {
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const fullDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  function shiftBg(val) {
    if (!val || val === '—') return '#F9FAFB';
    var v = String(val).toUpperCase();
    if (v.startsWith('DO')) return '#F3F4F6';
    if (v.startsWith('A/L')) return '#FEF3C7';
    return '#ECFDF5';
  }

  var weekTables = weekKeys.map(function(wk) {
    var shifts = allWeekShifts[wk] || {};
    var rows = days.map(function(d, i) {
      var val = shifts[d] || '—';
      return '<tr><td style="padding:8px 12px;border-bottom:1px solid #E5E7EB;font-weight:600;color:#374151;font-size:13px">' + fullDays[i] + '</td>' +
             '<td style="padding:8px 12px;border-bottom:1px solid #E5E7EB;background:' + shiftBg(val) + ';text-align:center;font-weight:600;color:#111827;font-size:13px">' + val + '</td></tr>';
    }).join('');

    return '<div style="margin-bottom:20px">' +
      '<h3 style="font-size:14px;color:#1E40AF;margin:0 0 8px;padding:8px 12px;background:#EFF6FF;border-radius:6px">Week of ' + wk + '</h3>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #E5E7EB;border-radius:6px;overflow:hidden">' +
        '<thead><tr style="background:#1E3A5F"><th style="padding:8px 12px;color:#fff;text-align:left;font-size:12px">Day</th><th style="padding:8px 12px;color:#fff;text-align:center;font-size:12px">Shift</th></tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +
    '</div>';
  }).join('');

  return '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E5E7EB">' +
    '<div style="background:linear-gradient(135deg,#1E3A5F,#1E40AF);padding:20px 24px;color:#fff">' +
      '<h2 style="margin:0;font-size:18px">Your Monthly Rota</h2>' +
      '<p style="margin:4px 0 0;opacity:0.8;font-size:14px">5 weeks starting ' + weekKeys[0] + '</p>' +
    '</div>' +
    '<div style="padding:20px 24px">' +
      '<p style="color:#374151;font-size:14px;margin:0 0 16px">Hi ' + staff.name + ',</p>' +
      '<p style="color:#374151;font-size:14px;margin:0 0 20px">Here are your shifts for the next 5 weeks:</p>' +
      weekTables +
      '<p style="color:#6B7280;font-size:13px;margin:8px 0 0">If you have any questions, please contact your manager.</p>' +
    '</div>' +
    '<div style="background:#F9FAFB;padding:14px 24px;border-top:1px solid #E5E7EB">' +
      '<p style="color:#9CA3AF;font-size:12px;margin:0;text-align:center">Waterfront Private Hospital — Rota Management System</p>' +
    '</div>' +
  '</div>';
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
  config.has_pin = isPinConfigured();
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

  // Handle admin PIN update
  if (config.admin_pin !== undefined) {
    let foundPin = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'admin_pin') {
        sheet.getRange(i + 1, 2).setValue(config.admin_pin);
        sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
        foundPin = true;
        break;
      }
    }
    if (!foundPin) sheet.appendRow(['admin_pin', config.admin_pin, new Date().toISOString()]);
  }

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
    sheet.appendRow(['id', 'name', 'group', 'department', 'hours', 'contractHours', 'email', 'phone', 'startDate', 'leaveEntitlement']);

    // Pre-populate with your current staff
    const staff = [
      // Admin
      ['adm-01', 'Janey', 'admin', '', 'FULL TIME', 37.5, '', '', '', 28],
      ['adm-02', 'Katrina', 'admin', '', 'FULL TIME', 37.5, '', '', '', 28],
      ['adm-03', 'Debbie', 'admin', '', 'PART TIME', 22.5, '', '', '', 28],
      ['adm-04', 'Danielle', 'admin', '', 'PART TIME', 15.5, '', '', '', 28],
      ['adm-05', 'Margaret', 'admin', '', 'PART TIME', 15, '', '', '', 28],
      ['adm-06', 'Lucy', 'admin', '', 'BANK', 0, '', '', '', 0],
      // Clinical
      ['cli-01', 'Morag', 'clinical', '', '', 37.5, '', '', '', 28],
      ['cli-02', 'Islay', 'clinical', '', '', 25, '', '', '', 28],
      ['cli-03', 'Oksana', 'clinical', '', '', 37.5, '', '', '', 28],
      ['cli-04', 'Christina', 'clinical', '', '', 37.5, '', '', '', 28],
      ['cli-05', 'Cezary', 'clinical', '', '', 37.5, '', '', '', 28],
      ['cli-06', 'Piel', 'clinical', '', '', 37.5, '', '', '', 28],
      ['cli-07', 'Inez', 'clinical', '', '', 30, '', '', '', 28],
      ['cli-08', 'Laura', 'clinical', '', '', 36, '', '', '', 28],
      ['cli-09', 'Kerry Ann', 'clinical', '', '', 33, '', '', '', 28],
      ['cli-10', 'Wojtek', 'clinical', '', '', 37.5, '', '', '', 28],
      // Bank
      ['bnk-01', 'Gleb', 'bank', '', '', 0, '', '', '', 0],
      ['bnk-02', 'Rona', 'bank', '', '', 0, '', '', '', 0],
      ['bnk-03', 'Sybil', 'bank', '', '', 0, '', '', '', 0],
      ['bnk-04', 'Gomes', 'bank', '', '', 0, '', '', '', 0],
      ['bnk-05', 'Myron', 'bank', '', '', 0, '', '', '', 0],
      ['bnk-06', 'Claudia', 'bank', 'OPD', '', 0, '', '', '', 0],
      ['bnk-07', 'Louise', 'bank', 'OPD', '', 0, '', '', '', 0],
      ['bnk-08', 'Sally', 'bank', 'OPD/Ward', '', 0, '', '', '', 0],
      ['bnk-09', 'Jo', 'bank', 'OPD', '', 0, '', '', '', 0],
      ['bnk-10', 'Eve', 'bank', 'Ward', '', 0, '', '', '', 0],
      ['bnk-11', 'Damien', 'bank', 'Ward', '', 0, '', '', '', 0],
      ['bnk-12', 'Draga', 'bank', 'Ward', '', 0, '', '', '', 0],
      ['bnk-13', 'Irina', 'bank', 'Ward', '', 0, '', '', '', 0],
      ['bnk-14', 'Olga', 'bank', 'ODP', '', 0, '', '', '', 0],
      ['bnk-15', 'Ashleigh', 'bank', 'ODP', '', 0, '', '', '', 0],
      ['bnk-16', 'Steve G', 'bank', 'ODP', '', 0, '', '', '', 0],
      ['bnk-17', 'Elisabeth', 'bank', 'TH', '', 0, '', '', '', 0],
      ['bnk-18', 'Clare', 'bank', 'Scrub/Rec', '', 0, '', '', '', 0],
      ['bnk-19', 'Isabella', 'bank', 'REC', '', 0, '', '', '', 0],
      ['bnk-20', 'Maxine', 'bank', 'ODP', '', 0, '', '', '', 0],
      ['bnk-21', 'Fiona B', 'bank', 'Scrub', '', 0, '', '', '', 0],
      ['bnk-22', 'Carl', 'bank', 'TH', '', 0, '', '', '', 0],
      ['bnk-23', 'Mariam', 'bank', '', '', 0, '', '', '', 0],
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
