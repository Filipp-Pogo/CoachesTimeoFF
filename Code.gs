// ============================================================
// WSC Coach Time Off & Coverage — Server-Side Logic
// ============================================================

var SPREADSHEET_ID = '15_Mr-VN5oF2pJwXRew1F4mMRO0xg3cZnaOUFDEKPHTs';
var APPROVALS_SHEET = 'Approvals Log';
var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbyfHxFTTF2KACoj2nT6n-KaY0E9mReUEwNJAJvNwoFtuGZ0orCcK6K8w-H6j2BcVodZ/exec';

var LUIS_EMAIL = 'Llopez@woodinvillesportsclub.com';
var FILIPP_EMAIL = 'fpogostkin@woodinvillesportsclub.com';

var COACHES = [
  { name: 'Gabriel Decamps',    email: 'gdecamps@woodinvillesportsclub.com' },
  { name: 'Rebeka Stolmar',     email: 'rstolmar@woodinvillesportsclub.com' },
  { name: 'Connor Vordale',     email: 'cvordale@woodinvillesportsclub.com' },
  { name: 'Ben Bethards',       email: 'bbethards@woodinvillesportsclub.com' },
  { name: 'Max Kan',            email: 'mkan@woodinvillesportsclub.com' },
  { name: 'Vivek Ramesh',       email: 'vramesh@woodinvillesportsclub.com' },
  { name: 'Kevin Hsieh',        email: 'khsieh@woodinvillesportsclub.com' },
  { name: 'Mitch Stewart',      email: 'mstewart@woodinvillesportsclub.com' },
  { name: 'Bond Minard',        email: 'bminard@woodinvillesportsclub.com' },
  { name: 'Maddy Bourguignon',  email: 'mbourguignon@woodinvillesportsclub.com' },
  { name: 'Maxim Groysman',     email: 'maximgroysman@gmail.com' },
  { name: 'Jon Bair',           email: 'jbair@woodinvillesportsclub.com' },
  { name: 'Filipp Pogostkin',   email: 'fpogostkin@woodinvillesportsclub.com' },
  { name: 'German Sanchez',     email: 'gsanchez@woodinvillesportsclub.com' },
  { name: 'Alaina Kim',         email: 'akim@woodinvillesportsclub.com' },
  { name: 'Oliver Wakeman',     email: 'owakeman@woodinvillesportsclub.com' },
  { name: 'Zach Brooks',        email: 'zbrooks@woodinvillesportsclub.com' },
  { name: 'Nirbhay Agarwal',    email: 'nagarwal@woodinvillesportsclub.com' },
  { name: 'Liam Tsinker',       email: 'ltsinker@woodinvillesportsclub.com' },
  { name: 'Stella Kim',         email: 'skim@woodinvillesportsclub.com' },
  { name: 'Daniel Jarvie',      email: 'djarvie@woodinvillesportsclub.com' },
  { name: 'Jackie S.',          email: 'jackies@woodinvillesportsclub.com' },
  { name: 'Luis Lopez',         email: 'llopez@woodinvillesportsclub.com' }
];

var CLASS_OPTIONS = [
  'Core Red', 'Core Orange', 'Core Green', 'Core Yellow',
  'Foundations', 'Junior Prep', 'Tier 1 HS', 'JASA', 'ASA', 'FTA',
  'Adult P&P', 'Adult T&L', 'Adult Small Group', 'Adult Intro',
  'Adult Team Practice', 'Adult Shot Spotlight'
];

var NOTICE_MULTI_DAY = 21;
var NOTICE_SHORT = 5;
var BLACKOUT_PERIODS = [];

// ============================================================
// WEB APP ROUTER
// ============================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'form';

  if (action === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('WSC Coach Time Off — Admin Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (action === 'approve' && e.parameter.id) {
    return handleLegacyApprove_(e.parameter.id);
  }

  if (action === 'deny' && e.parameter.id) {
    return handleLegacyDenyForm_(e.parameter.id);
  }

  if (action === 'deny_confirm' && e.parameter.id) {
    var reason = e.parameter.reason || 'No reason provided';
    return handleLegacyDenyConfirm_(e.parameter.id, reason);
  }

  // Default: coach form
  return HtmlService.createTemplateFromFile('CoachForm')
    .evaluate()
    .setTitle('WSC Coach Time Off Request')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// DATA ACCESSORS (called from client via google.script.run)
// ============================================================

function getCoaches() {
  return COACHES;
}

function getClassOptions() {
  return CLASS_OPTIONS;
}

var DASHBOARD_PASSWORD = 'Tier1tennis2026';

function validateDashboardAccess(password) {
  return password === DASHBOARD_PASSWORD;
}

function getLoggedInCoach() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    for (var i = 0; i < COACHES.length; i++) {
      if (COACHES[i].email.toLowerCase() === email) {
        return { name: COACHES[i].name, email: COACHES[i].email };
      }
    }
  } catch (e) {}
  return null;
}

// ============================================================
// SUBMISSION PROCESSING
// ============================================================

function processSubmission(formData) {
  var responseId = Utilities.getUuid();
  var now = new Date();
  var startDate = new Date(formData.startDate + 'T00:00:00');
  var endDate = new Date(formData.endDate + 'T00:00:00');

  // Resolve sub coach emails from roster
  var classes = formData.classes || [];
  for (var i = 0; i < classes.length; i++) {
    if (classes[i].subName && classes[i].subName !== 'Other (not listed)') {
      var match = findCoachByName_(classes[i].subName);
      if (match) classes[i].subEmail = match.email;
    }
  }

  // Check auto-disqualification
  var disqualification = checkAutoDisqualification_(startDate, endDate, now);
  if (disqualification) {
    appendToApprovalsLog_(now, responseId, formData.coachName, formData.coachEmail, startDate, endDate, 'AUTO-DENIED', disqualification.reason, formData.reason, classes);
    sendAutoDenialToCoach_(formData, disqualification.reason);
    sendAutoDenialToLuis_(formData, disqualification.reason);
    return { success: false, autoDenied: true, reason: disqualification.reason };
  }

  // Check sub conflicts (blocking — auto-deny if any sub has a conflict)
  var conflicts = checkSubConflicts_(classes, startDate, endDate);
  if (conflicts.length > 0) {
    var conflictReason = 'Coverage conflict: ' + conflicts.join('; ');
    appendToApprovalsLog_(now, responseId, formData.coachName, formData.coachEmail, startDate, endDate, 'AUTO-DENIED', conflictReason, formData.reason, classes);
    sendAutoDenialToCoach_(formData, conflictReason);
    sendAutoDenialToLuis_(formData, conflictReason);
    return { success: false, autoDenied: true, reason: conflictReason };
  }

  // Write to Approvals Log
  appendToApprovalsLog_(now, responseId, formData.coachName, formData.coachEmail, startDate, endDate, 'PENDING', '', formData.reason, classes);

  // Store in PropertiesService
  var stored = {
    coachName: formData.coachName,
    coachEmail: formData.coachEmail,
    classes: classes,
    startDate: formData.startDate,
    endDate: formData.endDate,
    reason: formData.reason,
    submittedAt: now.toISOString(),
    reminderCount: 0,
    lastReminderAt: null,
    filippEscalated: false
  };
  PropertiesService.getScriptProperties().setProperty('response_' + responseId, JSON.stringify(stored));

  // Send emails
  sendSubNotifications_(formData, classes);
  sendCoachConfirmation_(formData);
  sendApprovalRequestToLuis_(formData, responseId, classes, conflicts);

  return { success: true, responseId: responseId };
}

// ============================================================
// AUTO-DISQUALIFICATION
// ============================================================

function checkAutoDisqualification_(startDate, endDate, now) {
  var absenceDays = Math.ceil((endDate - startDate) / 86400000) + 1;

  // Check blackout periods
  for (var i = 0; i < BLACKOUT_PERIODS.length; i++) {
    var bp = BLACKOUT_PERIODS[i];
    var bpStart = new Date(bp.start + 'T00:00:00');
    var bpEnd = new Date(bp.end + 'T23:59:59');
    if (startDate <= bpEnd && endDate >= bpStart) {
      return { reason: 'Dates overlap blackout period: ' + bp.label + ' (' + bp.start + ' to ' + bp.end + ')' };
    }
  }

  if (absenceDays >= 3) {
    var calendarDays = Math.ceil((startDate - now) / 86400000);
    if (calendarDays < NOTICE_MULTI_DAY) {
      return { reason: 'Insufficient notice for 3+ day absence. Required: ' + NOTICE_MULTI_DAY + ' calendar days. Provided: ' + calendarDays + ' days.' };
    }
  } else {
    var bizDays = countBusinessDays_(now, startDate);
    if (bizDays < NOTICE_SHORT) {
      return { reason: 'Insufficient notice for 1-2 day absence. Required: ' + NOTICE_SHORT + ' business days. Provided: ' + bizDays + ' business days.' };
    }
  }

  return null;
}

function countBusinessDays_(from, to) {
  var count = 0;
  var d = new Date(from);
  d.setHours(0, 0, 0, 0);
  var target = new Date(to);
  target.setHours(0, 0, 0, 0);
  while (d < target) {
    d.setDate(d.getDate() + 1);
    var day = d.getDay();
    if (day !== 0 && day !== 6) count++;
  }
  return count;
}

// ============================================================
// SUB CONFLICT CHECK
// ============================================================

function checkSubConflicts_(classes, startDate, endDate) {
  var props = PropertiesService.getScriptProperties().getProperties();
  var approvedAbsences = getApprovedAbsences_();
  var conflicts = [];

  for (var i = 0; i < classes.length; i++) {
    var subName = classes[i].subName;
    if (!subName || subName === 'Other (not listed)') continue;

    // Check pending requests in PropertiesService
    for (var key in props) {
      if (key.indexOf('response_') !== 0) continue;
      try {
        var req = JSON.parse(props[key]);
        if (req.coachName === subName) {
          var reqStart = new Date(req.startDate + 'T00:00:00');
          var reqEnd = new Date(req.endDate + 'T00:00:00');
          if (startDate <= reqEnd && endDate >= reqStart) {
            conflicts.push(subName + ' has a pending time-off request from ' + req.startDate + ' to ' + req.endDate);
          }
        }
      } catch (e) {}
    }

    // Check approved requests in Approvals Log
    for (var a = 0; a < approvedAbsences.length; a++) {
      var ab = approvedAbsences[a];
      if (ab.coachName === subName) {
        if (startDate <= ab.endDate && endDate >= ab.startDate) {
          conflicts.push(subName + ' has an approved time-off from ' + formatDate_(ab.startDate) + ' to ' + formatDate_(ab.endDate));
        }
      }
    }
  }
  return conflicts;
}

// Client-callable version for real-time warnings
function checkSubConflictsForClient(subName, startDateStr, endDateStr) {
  if (!subName || subName === 'Other (not listed)') return [];
  var startDate = new Date(startDateStr + 'T00:00:00');
  var endDate = new Date(endDateStr + 'T00:00:00');
  var props = PropertiesService.getScriptProperties().getProperties();
  var approvedAbsences = getApprovedAbsences_();
  var conflicts = [];

  // Check pending requests
  for (var key in props) {
    if (key.indexOf('response_') !== 0) continue;
    try {
      var req = JSON.parse(props[key]);
      if (req.coachName === subName) {
        var reqStart = new Date(req.startDate + 'T00:00:00');
        var reqEnd = new Date(req.endDate + 'T00:00:00');
        if (startDate <= reqEnd && endDate >= reqStart) {
          conflicts.push(subName + ' has a pending time-off request from ' + req.startDate + ' to ' + req.endDate);
        }
      }
    } catch (e) {}
  }

  // Check approved requests
  for (var a = 0; a < approvedAbsences.length; a++) {
    var ab = approvedAbsences[a];
    if (ab.coachName === subName) {
      if (startDate <= ab.endDate && endDate >= ab.startDate) {
        conflicts.push(subName + ' has an approved time-off from ' + formatDate_(ab.startDate) + ' to ' + formatDate_(ab.endDate));
      }
    }
  }
  return conflicts;
}

function getApprovedAbsences_() {
  var sheet = getOrCreateApprovalsSheet_();
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][8]) === 'APPROVED') {
      results.push({
        coachName: String(data[i][2]),
        startDate: new Date(data[i][4]),
        endDate: new Date(data[i][5])
      });
    }
  }
  return results;
}

// ============================================================
// APPROVALS LOG
// ============================================================

function getOrCreateApprovalsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(APPROVALS_SHEET);
  if (sheet) {
    // Check if headers match the new 10-column format
    var firstRow = sheet.getRange(1, 1, 1, 10).getValues()[0];
    if (String(firstRow[6]) !== 'Reason') {
      // Old format — rename old sheet and create fresh one
      sheet.setName('Approvals Log (old)');
      sheet = null;
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(APPROVALS_SHEET);
    sheet.appendRow(['Timestamp', 'Response ID', 'Coach Name', 'Coach Email', 'Start Date', 'End Date', 'Reason', 'Classes/Coverage', 'Status', 'Notes']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(8, 350);
    sheet.setColumnWidth(9, 100);
    sheet.setColumnWidth(10, 250);
  }
  return sheet;
}

function formatClassesCoverage_(classes) {
  if (!classes || classes.length === 0) return '';
  var lines = [];
  for (var i = 0; i < classes.length; i++) {
    var c = classes[i];
    var sub = c.subName || 'No sub assigned';
    lines.push(c.className + ' → ' + sub);
  }
  return lines.join(', ');
}

function appendToApprovalsLog_(timestamp, responseId, coachName, coachEmail, startDate, endDate, status, notes, reason, classes) {
  try {
    var sheet = getOrCreateApprovalsSheet_();
    var coverageStr = formatClassesCoverage_(classes);
    sheet.appendRow([timestamp, responseId, coachName, coachEmail, startDate, endDate, reason || '', coverageStr, status, notes]);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('ERROR writing to Approvals Log: ' + e.message);
    throw e;
  }
}

function updateApprovalLog(responseId, status, notes) {
  var sheet = getOrCreateApprovalsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(responseId)) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 10).setValue(notes);
      return true;
    }
  }
  return false;
}

// ============================================================
// DATE UTILITIES
// ============================================================

function getClassDates(startDateStr, endDateStr, dayTimeStr) {
  var startDate = new Date(startDateStr + 'T00:00:00');
  var endDate = new Date(endDateStr + 'T00:00:00');
  var dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

  // Parse multiple day entries (e.g. "Monday 4pm, Wednesday 4pm")
  var parts = dayTimeStr.split(',');
  var results = [];

  for (var p = 0; p < parts.length; p++) {
    var part = parts[p].trim();
    var firstWord = part.split(/\s+/)[0].toLowerCase();
    var targetDay = -1;
    for (var d = 0; d < dayNames.length; d++) {
      if (dayNames[d].indexOf(firstWord) === 0) {
        targetDay = d;
        break;
      }
    }
    if (targetDay === -1) continue;

    var current = new Date(startDate);
    while (current <= endDate) {
      if (current.getDay() === targetDay) {
        results.push(formatDateShort_(current) + ' — ' + part);
      }
      current.setDate(current.getDate() + 1);
    }
  }
  return results;
}

function formatDateShort_(date) {
  var days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return days[date.getDay()] + ', ' + months[date.getMonth()] + ' ' + date.getDate();
}

function formatDate_(date) {
  if (!(date instanceof Date)) date = new Date(date);
  return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
}

// ============================================================
// COACH LOOKUP
// ============================================================

function findCoachByName_(name) {
  for (var i = 0; i < COACHES.length; i++) {
    if (COACHES[i].name === name) return COACHES[i];
  }
  return null;
}

// ============================================================
// EMAIL FUNCTIONS
// ============================================================

var EMAIL_FOOTER = '\n\n━━━━━━━━━━━━━━━━━\nWoodinville Sports Club — Automated Notification';

function sendSubNotifications_(formData, classes) {
  for (var i = 0; i < classes.length; i++) {
    var c = classes[i];
    if (!c.subEmail || !c.subName) continue;
    var subject = 'WSC Coverage Request — ' + c.className;
    var body = 'Hi ' + c.subName.split(' ')[0] + ',\n\n' +
      formData.coachName + ' has requested time off and listed you as a substitute coach.\n\n' +
      '━━━━━━━━━━━━━━━━━\n' +
      'CLASS: ' + c.className + '\n' +
      'DAY & TIME: ' + c.classTime + '\n' +
      'ABSENCE PERIOD: ' + formData.startDate + ' to ' + formData.endDate + '\n' +
      '━━━━━━━━━━━━━━━━━\n\n' +
      'Please confirm with ' + formData.coachName + ' that you are available to cover.\n' +
      'If you have any questions, contact Luis Lopez.' +
      EMAIL_FOOTER;
    MailApp.sendEmail(c.subEmail, subject, body);
  }
}

function sendCoachConfirmation_(formData) {
  var classes = formData.classes || [];
  var subject = 'WSC Time Off Request Received';
  var body = 'Hi ' + formData.coachName.split(' ')[0] + ',\n\n' +
    'Your time off request has been received and is pending approval.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'DATES: ' + formData.startDate + ' to ' + formData.endDate + '\n' +
    'REASON: ' + formData.reason + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n';

  if (classes.length > 0) {
    body += 'YOUR COVERAGE PLAN:\n';
    for (var i = 0; i < classes.length; i++) {
      var c = classes[i];
      body += '  - ' + c.className + ' (' + (c.classTime || 'TBD') + ') — Sub: ' + c.subName + (c.subEmail && c.subName === 'Other (not listed)' ? ' (' + c.subEmail + ')' : '') + '\n';
    }
    body += '\nPlease double-check the coverage details above. If anything is wrong, contact Luis Lopez.\n\n';
  }

  body += 'REMINDERS:\n' +
    '- 3+ day requests require 3 weeks advance notice\n' +
    '- 1-2 day requests require 5 business days advance notice\n' +
    '- Emergency situations: notify leadership ASAP\n' +
    '- Sick leave must be reported through Gusto\n\n' +
    'You will receive an email once your request is approved or denied.' +
    EMAIL_FOOTER;
  MailApp.sendEmail(formData.coachEmail, subject, body);
}

function sendApprovalRequestToLuis_(formData, responseId, classes, conflicts) {
  var subject = 'ACTION REQUIRED — Time Off Request from ' + formData.coachName;

  var body = 'A new time off request needs your review.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + formData.coachName + ' (' + formData.coachEmail + ')\n' +
    'DATES: ' + formData.startDate + ' to ' + formData.endDate + '\n' +
    'REASON: ' + formData.reason + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n';

  body += 'COVERAGE PLAN:\n';
  for (var i = 0; i < classes.length; i++) {
    var c = classes[i];
    body += '\n  Class: ' + c.className + '\n';
    body += '  Day & Time: ' + c.classTime + '\n';
    body += '  Sub Coach: ' + c.subName + ' (' + (c.subEmail || 'no email') + ')\n';

    var dates = getClassDates(formData.startDate, formData.endDate, c.classTime);
    if (dates.length > 0) {
      body += '  Specific dates:\n';
      for (var d = 0; d < dates.length; d++) {
        body += '    - ' + dates[d] + '\n';
      }
    }
  }

  if (conflicts.length > 0) {
    body += '\n⚠ CONFLICT WARNINGS:\n';
    for (var j = 0; j < conflicts.length; j++) {
      body += '  - ' + conflicts[j] + '\n';
    }
  }

  body += '\n━━━━━━━━━━━━━━━━━\n';
  body += 'APPROVE: ' + WEBAPP_URL + '?action=approve&id=' + responseId + '\n';
  body += 'DENY: ' + WEBAPP_URL + '?action=deny&id=' + responseId + '\n';
  body += '\nOr use the dashboard: ' + WEBAPP_URL + '?action=dashboard';
  body += EMAIL_FOOTER;

  MailApp.sendEmail(LUIS_EMAIL, subject, body);
}

function sendAutoDenialToCoach_(formData, reason) {
  var subject = 'WSC Time Off Request — Automatically Declined';
  var body = 'Hi ' + formData.coachName.split(' ')[0] + ',\n\n' +
    'Your time off request has been automatically declined.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'DATES: ' + formData.startDate + ' to ' + formData.endDate + '\n' +
    'REASON FOR DENIAL: ' + reason + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n' +
    'POLICY REMINDERS:\n' +
    '- 3+ day requests require 3 weeks advance notice\n' +
    '- 1-2 day requests require 5 business days advance notice\n' +
    '- Blackout dates: First week of each season, National Championships, WSC major events\n' +
    '- Emergency situations: notify leadership ASAP\n\n' +
    'If this is an emergency, please contact Luis Lopez directly.' +
    EMAIL_FOOTER;
  MailApp.sendEmail(formData.coachEmail, subject, body);
}

function sendAutoDenialToLuis_(formData, reason) {
  var subject = 'FYI — Auto-Denied Time Off: ' + formData.coachName;
  var body = 'A time off request was automatically denied.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + formData.coachName + ' (' + formData.coachEmail + ')\n' +
    'DATES: ' + formData.startDate + ' to ' + formData.endDate + '\n' +
    'REASON FOR REQUEST: ' + formData.reason + '\n' +
    'DENIAL REASON: ' + reason + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n' +
    'No action required unless this is an emergency override.' +
    EMAIL_FOOTER;
  MailApp.sendEmail(LUIS_EMAIL, subject, body);
}

function sendApprovalToCoach_(data) {
  var subject = 'WSC Time Off Request — APPROVED';
  var body = 'Hi ' + data.coachName.split(' ')[0] + ',\n\n' +
    'Your time off request has been APPROVED.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'DATES: ' + data.startDate + ' to ' + data.endDate + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n';

  if (data.classes && data.classes.length > 0) {
    body += 'COVERAGE SUMMARY:\n';
    for (var i = 0; i < data.classes.length; i++) {
      var c = data.classes[i];
      body += '  - ' + c.className + ' (' + c.classTime + ') — Sub: ' + c.subName + '\n';
    }
  }

  body += '\nPlease coordinate any handoff details with your substitute coaches.' +
    EMAIL_FOOTER;
  MailApp.sendEmail(data.coachEmail, subject, body);
}

function sendCourtReserveUpdateToLuis_(data) {
  var subject = 'Court Reserve Update Required — ' + data.coachName + ' Time Off';
  var submittedLabel = data.submittedAt ? 'SUBMITTED: ' + formatDate_(new Date(data.submittedAt)) + '\n' : '';
  var body = 'The following time off request has been APPROVED. Please update Court Reserve.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + data.coachName + '\n' +
    'ABSENCE: ' + data.startDate + ' to ' + data.endDate + '\n' +
    submittedLabel +
    '━━━━━━━━━━━━━━━━━\n\n' +
    'COURT RESERVE CHANGES NEEDED:\n';

  for (var i = 0; i < data.classes.length; i++) {
    var c = data.classes[i];
    body += '\n  Class: ' + c.className + '\n';
    body += '  Remove: ' + data.coachName + '\n';
    body += '  Add: ' + c.subName + (c.subEmail ? ' (' + c.subEmail + ')' : '') + '\n';

    var dates = getClassDates(data.startDate, data.endDate, c.classTime);
    if (dates.length > 0) {
      body += '  Dates to update:\n';
      for (var d = 0; d < dates.length; d++) {
        body += '    - ' + dates[d] + '\n';
      }
    }
  }

  body += EMAIL_FOOTER;
  MailApp.sendEmail(LUIS_EMAIL, subject, body);
}

function sendDenialToCoach_(data, reason) {
  var subject = 'WSC Time Off Request — DENIED';
  var body = 'Hi ' + data.coachName.split(' ')[0] + ',\n\n' +
    'Your time off request has been denied.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'DATES: ' + data.startDate + ' to ' + data.endDate + '\n' +
    'REASON: ' + reason + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n' +
    'If you have questions, please contact Luis Lopez.' +
    EMAIL_FOOTER;
  MailApp.sendEmail(data.coachEmail, subject, body);
}

function sendDenialSummaryToLuis_(data, reason) {
  var subject = 'Confirmed — Denied Time Off: ' + data.coachName;
  var body = 'You denied the following time off request.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + data.coachName + '\n' +
    'DATES: ' + data.startDate + ' to ' + data.endDate + '\n' +
    'DENIAL REASON: ' + reason + '\n' +
    '━━━━━━━━━━━━━━━━━' +
    EMAIL_FOOTER;
  MailApp.sendEmail(LUIS_EMAIL, subject, body);
}

// ============================================================
// DASHBOARD ACTIONS (called from Dashboard.html)
// ============================================================

function getPendingRequests() {
  var props = PropertiesService.getScriptProperties().getProperties();

  // Load sheet statuses in one call to avoid N+1 spreadsheet opens
  var statusMap = getAllStatuses_();

  var pending = [];
  for (var key in props) {
    if (key.indexOf('response_') !== 0) continue;
    try {
      var data = JSON.parse(props[key]);
      var id = key.replace('response_', '');
      if (statusMap[id] === 'PENDING') {
        data.responseId = id;
        data.pendingHours = Math.round((new Date() - new Date(data.submittedAt)) / 3600000);
        pending.push(data);
      }
    } catch (e) {}
  }
  pending.sort(function(a, b) { return new Date(a.submittedAt) - new Date(b.submittedAt); });
  return pending;
}

function getAllStatuses_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(APPROVALS_SHEET);
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[String(data[i][1])] = String(data[i][6]);
  }
  return map;
}

function getRecentHistory() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(APPROVALS_SHEET);
  var data = sheet.getDataRange().getValues();
  var history = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][6]);
    if (status !== 'PENDING') {
      history.push({
        timestamp: data[i][0],
        responseId: data[i][1],
        coachName: data[i][2],
        coachEmail: data[i][3],
        startDate: data[i][4],
        endDate: data[i][5],
        status: status,
        notes: data[i][7]
      });
    }
    if (history.length >= 30) break;
  }
  return history;
}

function getStatusFromSheet_(responseId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(APPROVALS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(responseId)) {
      return String(data[i][6]);
    }
  }
  return null;
}

function approveRequest(responseId) {
  var propKey = 'response_' + responseId;
  var raw = PropertiesService.getScriptProperties().getProperty(propKey);
  if (!raw) return { success: false, error: 'Request not found or already processed' };

  // Check sheet status to prevent double-processing
  var currentStatus = getStatusFromSheet_(responseId);
  if (currentStatus !== 'PENDING') {
    PropertiesService.getScriptProperties().deleteProperty(propKey);
    return { success: false, error: 'Request already ' + (currentStatus || 'processed') };
  }

  var data = JSON.parse(raw);
  updateApprovalLog(responseId, 'APPROVED', '');
  sendApprovalToCoach_(data);
  sendCourtReserveUpdateToLuis_(data);
  PropertiesService.getScriptProperties().deleteProperty(propKey);

  return { success: true };
}

function denyRequest(responseId, reason) {
  var propKey = 'response_' + responseId;
  var raw = PropertiesService.getScriptProperties().getProperty(propKey);
  if (!raw) return { success: false, error: 'Request not found or already processed' };

  // Check sheet status to prevent double-processing
  var currentStatus = getStatusFromSheet_(responseId);
  if (currentStatus !== 'PENDING') {
    PropertiesService.getScriptProperties().deleteProperty(propKey);
    return { success: false, error: 'Request already ' + (currentStatus || 'processed') };
  }

  var data = JSON.parse(raw);
  updateApprovalLog(responseId, 'DENIED', reason);
  sendDenialToCoach_(data, reason);
  sendDenialSummaryToLuis_(data, reason);
  PropertiesService.getScriptProperties().deleteProperty(propKey);

  return { success: true };
}

// ============================================================
// LEGACY EMAIL LINK HANDLERS
// ============================================================

function handleLegacyApprove_(responseId) {
  var result = approveRequest(responseId);
  var html;
  if (result.success) {
    html = '<html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
      '<h2 style="color:#1a5632;">Request Approved</h2>' +
      '<p>The coach has been notified and a Court Reserve update email has been sent.</p>' +
      '<p><a href="' + WEBAPP_URL + '?action=dashboard">Go to Dashboard</a></p></body></html>';
  } else {
    html = '<html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
      '<h2 style="color:#c0392b;">Error</h2><p>' + result.error + '</p></body></html>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('WSC — Approval')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function handleLegacyDenyForm_(responseId) {
  var html = '<html><body style="font-family:sans-serif;max-width:500px;margin:40px auto;padding:20px;">' +
    '<h2 style="color:#c0392b;">Deny Request</h2>' +
    '<form action="' + WEBAPP_URL + '" method="get">' +
    '<input type="hidden" name="action" value="deny_confirm">' +
    '<input type="hidden" name="id" value="' + responseId + '">' +
    '<label style="display:block;margin-bottom:8px;font-weight:bold;">Reason for denial:</label>' +
    '<textarea name="reason" rows="4" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;" required></textarea>' +
    '<button type="submit" style="margin-top:12px;background:#c0392b;color:white;border:none;padding:10px 24px;border-radius:4px;cursor:pointer;">Confirm Denial</button>' +
    '</form></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('WSC — Deny Request')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function handleLegacyDenyConfirm_(responseId, reason) {
  var result = denyRequest(responseId, reason);
  var html;
  if (result.success) {
    html = '<html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
      '<h2 style="color:#c0392b;">Request Denied</h2>' +
      '<p>The coach has been notified.</p>' +
      '<p><a href="' + WEBAPP_URL + '?action=dashboard">Go to Dashboard</a></p></body></html>';
  } else {
    html = '<html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
      '<h2 style="color:#c0392b;">Error</h2><p>' + result.error + '</p></body></html>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('WSC — Denial Confirmed')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// PENDING REQUEST REMINDERS & ESCALATION (daily trigger)
// ============================================================

function checkPendingRequests() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = new Date();
  var statusMap = getAllStatuses_();

  for (var key in all) {
    if (key.indexOf('response_') !== 0) continue;

    var responseId = key.replace('response_', '');
    var data;
    try { data = JSON.parse(all[key]); } catch (e) { continue; }

    // Verify still PENDING
    if (statusMap[responseId] !== 'PENDING') {
      props.deleteProperty(key);
      continue;
    }

    var submitted = new Date(data.submittedAt);
    var hoursElapsed = (now - submitted) / 3600000;

    if (hoursElapsed < 48) continue;

    if (hoursElapsed >= 168 && !data.filippEscalated) {
      // 7+ days — escalate to Filipp
      sendEscalationToFilipp_(data, responseId);
      data.filippEscalated = true;
      data.reminderCount = (data.reminderCount || 0) + 1;
      data.lastReminderAt = now.toISOString();
      props.setProperty(key, JSON.stringify(data));
      continue;
    }

    // 48h to 7d — remind Luis (max once per 23 hours)
    if (data.lastReminderAt) {
      var lastReminder = new Date(data.lastReminderAt);
      if ((now - lastReminder) / 3600000 < 23) continue;
    }

    sendPendingReminderToLuis_(data, responseId, hoursElapsed);
    data.reminderCount = (data.reminderCount || 0) + 1;
    data.lastReminderAt = now.toISOString();
    props.setProperty(key, JSON.stringify(data));
  }
}

function sendPendingReminderToLuis_(data, responseId, hoursElapsed) {
  var days = Math.round(hoursElapsed / 24);
  var subject = 'REMINDER — Pending Time Off Request (' + days + ' days) — ' + data.coachName;
  var body = 'This time off request is still pending your review.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + data.coachName + '\n' +
    'DATES: ' + data.startDate + ' to ' + data.endDate + '\n' +
    'REASON: ' + data.reason + '\n' +
    'SUBMITTED: ' + data.submittedAt + '\n' +
    'PENDING FOR: ' + days + ' days\n' +
    '━━━━━━━━━━━━━━━━━\n\n' +
    'APPROVE: ' + WEBAPP_URL + '?action=approve&id=' + responseId + '\n' +
    'DENY: ' + WEBAPP_URL + '?action=deny&id=' + responseId + '\n' +
    '\nDashboard: ' + WEBAPP_URL + '?action=dashboard' +
    EMAIL_FOOTER;
  MailApp.sendEmail(LUIS_EMAIL, subject, body);
}

function sendEscalationToFilipp_(data, responseId) {
  var subject = 'ESCALATION — Time Off Request Pending 7+ Days — ' + data.coachName;
  var body = 'The following time off request has been pending for over 7 days without action.\n\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    'COACH: ' + data.coachName + ' (' + data.coachEmail + ')\n' +
    'DATES: ' + data.startDate + ' to ' + data.endDate + '\n' +
    'REASON: ' + data.reason + '\n' +
    'SUBMITTED: ' + data.submittedAt + '\n' +
    '━━━━━━━━━━━━━━━━━\n\n';

  if (data.classes && data.classes.length > 0) {
    body += 'COVERAGE PLAN:\n';
    for (var i = 0; i < data.classes.length; i++) {
      var c = data.classes[i];
      body += '  - ' + c.className + ' (' + c.classTime + ') — Sub: ' + c.subName + '\n';
    }
    body += '\n';
  }

  body += 'APPROVE: ' + WEBAPP_URL + '?action=approve&id=' + responseId + '\n';
  body += 'DENY: ' + WEBAPP_URL + '?action=deny&id=' + responseId + '\n';
  body += '\nDashboard: ' + WEBAPP_URL + '?action=dashboard' +
    EMAIL_FOOTER;
  MailApp.sendEmail(FILIPP_EMAIL, subject, body);
}

// ============================================================
// TRIGGER SETUP — Run once after deploying
// ============================================================

function setupTrigger() {
  // Delete existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // Daily pending request check
  ScriptApp.newTrigger('checkPendingRequests')
    .timeBased()
    .everyHours(24)
    .create();

  Logger.log('Triggers set up successfully.');
}
