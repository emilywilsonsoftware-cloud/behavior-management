// ============================================================
// Behavior Tracker — Google Apps Script Web App
// Setup.gs  v4 — Split first/last/middle name columns
//
// CHANGES FROM v3:
//   • Students sheet: StudentName split into FirstName, LastName, MiddleName
//     New layout: StudentID | FirstName | LastName | MiddleName | Grade |
//                 CurrentPoints | PointsLastUpdated
//   • Staff sheet: StaffName split into FirstName, LastName
//     New layout: FirstName | LastName | StaffEmail | Role
//   • ParentContacts sheet: ParentName split into FirstName, LastName
//     New layout: StudentID | ParentGUID | FirstName | LastName |
//                 ParentEmail | Phone
// ============================================================

// ── Sheet name constants (must match Code.gs) ─────────────────
var SS_REFERRALS   = 'Referrals';
var SS_STUDENTS    = 'Students';
var SS_STAFF       = 'Staff';
var SS_PARENT      = 'ParentContacts';
var SS_CONFIG      = 'Config';
var SS_INFRACTIONS = 'Infractions';

// =============================================================
// SPREADSHEET MENU
// =============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BehaviorTracker Admin')
    .addItem('Run Initial Setup',          'setupSpreadsheet')
    .addSeparator()
    .addItem('Setup Daily Email Trigger',  'setupDailyTrigger')
    .addItem('Remove Daily Email Trigger', 'removeDailyTrigger')
    .addSeparator()
    .addItem('Reset Semester Points (NEW SEMESTER)', 'resetSemesterPointsMenu')
    .addToUi();
}

// =============================================================
// INITIAL SETUP
// =============================================================

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  _setupConfig(ss);
  _setupInfractions(ss);
  _setupReferrals(ss);
  _setupStudents(ss);
  _setupStaff(ss);
  _setupParentContacts(ss);

  SpreadsheetApp.flush();
  ui.alert(
    'Setup Complete',
    'All sheets have been created.\n\n' +
    'Next steps:\n' +
    '1. Update the Config sheet with your school name and settings.\n' +
    '2. Replace sample data in Students, Staff, and ParentContacts.\n' +
    '3. Add at least one admin row to the Staff sheet (Role = admin).\n' +
    '4. Deploy the web app and run setupDailyTrigger from this menu.',
    ui.ButtonSet.OK
  );
}

// ── Config sheet ──────────────────────────────────────────────
function _setupConfig(ss) {
  if (ss.getSheetByName(SS_CONFIG)) return;
  var sheet = ss.insertSheet(SS_CONFIG);

  // Redirections and Motivations columns intentionally removed —
  // Behavior Context section no longer exists on the referral form.
  var headers = [
    'SchoolName', 'SemesterStartPoints', 'EmailNotificationsEnabled',
    'DailyEmailSendTime', 'EmailFooterText',
    'Locations', 'NineWeeksStartDates'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var row2 = [
    'Our School', '100', 'Yes', '15:30',
    'If you have questions, please contact the school office.\n\n' +
    'This is an automated message — please do not reply.',
    'Classroom', ''
  ];
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);

  var extraLocations = ['Hallway', 'Cafeteria', 'Gym', 'Bus', 'Restroom', 'Office'];
  var extraNineWeeks = ['2024-08-01', '2024-10-15', '2025-01-08', '2025-03-18'];
  var maxExtra = Math.max(extraLocations.length, extraNineWeeks.length);
  for (var r = 0; r < maxExtra; r++) {
    sheet.getRange(r + 3, 1, 1, 7).setValues([[
      '', '', '', '', '',
      extraLocations[r] || '',
      extraNineWeeks[r] || ''
    ]]);
  }
  sheet.setColumnWidth(5, 300);
  sheet.autoResizeColumns(1, 4);
}

// ── Infractions sheet ─────────────────────────────────────────
function _setupInfractions(ss) {
  if (ss.getSheetByName(SS_INFRACTIONS)) return;
  var sheet = ss.insertSheet(SS_INFRACTIONS);

  var headers = ['InfractionName', 'PointValue', 'Severity', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var data = [
    ['Dress Code Violation',         -3,  'Minor',    ''],
    ['Failure to Follow Directions', -5,  'Minor',    ''],
    ['Disruptive Behavior',          -5,  'Minor',    ''],
    ['Disrespect',                   -7,  'Minor',    ''],
    ['Bullying',                     -15, 'Major',    ''],
    ['Fighting',                     -20, 'Major',    ''],
    ['Vandalism',                    -15, 'Major',    ''],
    ['Harassment',                   -15, 'Major',    ''],
    ['Technology Violation',         -5,  'Minor',    ''],
    ['Tardy',                        -2,  'Minor',    ''],
    ['Outstanding Behavior',          5,  'Positive', ''],
    ['Helping Others',                3,  'Positive', ''],
    ['Academic Achievement',          5,  'Positive', ''],
    ['Perfect Attendance',            3,  'Positive', '']
  ];
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.autoResizeColumns(1, headers.length);
}

// ── Referrals sheet ───────────────────────────────────────────
// StudentName stored as "First Last" at submit time for historical record.
function _setupReferrals(ss) {
  if (ss.getSheetByName(SS_REFERRALS)) return;
  var sheet = ss.insertSheet(SS_REFERRALS);
  var headers = [
    'ID', 'Timestamp', 'StudentID', 'StudentName', 'Grade',
    'IncidentDate', 'IncidentTime', 'Location', 'InfractionType',
    'Severity', 'PointValue', 'PointsBeforeReferral', 'PointsAfterReferral',
    'Description', 'IncludeDescriptionInEmail',
    'TeacherName', 'TeacherEmail',
    'ParentNotified', 'TeacherNotified', 'Status', 'AdminNotes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, headers.length);
}

// ── Students sheet ────────────────────────────────────────────
// Layout: StudentID | FirstName | LastName | MiddleName | Grade |
//         CurrentPoints | PointsLastUpdated
function _setupStudents(ss) {
  if (ss.getSheetByName(SS_STUDENTS)) return;
  var sheet = ss.insertSheet(SS_STUDENTS);

  var headers = [
    'StudentID', 'FirstName', 'LastName', 'MiddleName',
    'Grade', 'CurrentPoints', 'PointsLastUpdated'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var data = [
    ['S001', 'Alice',  'Smith',    '', '6', 100, ''],
    ['S002', 'Bob',    'Jones',    '', '6', 100, ''],
    ['S003', 'Frank',  'Taylor',   '', '7', 100, ''],
    ['S004', 'Grace',  'Anderson', '', '7', 100, ''],
    ['S005', 'Henry',  'Brown',    '', '8', 100, ''],
    ['S006', 'Ivy',    'Davis',    '', '8', 100, '']
  ];
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.autoResizeColumns(1, headers.length);
}

// ── Staff sheet ───────────────────────────────────────────────
// Layout: FirstName | LastName | StaffEmail | Role
function _setupStaff(ss) {
  if (ss.getSheetByName(SS_STAFF)) return;
  var sheet = ss.insertSheet(SS_STAFF);

  var headers = ['FirstName', 'LastName', 'StaffEmail', 'Role'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var data = [
    ['Admin',  'User',  'admin@yourschool.edu',       'admin'  ],
    ['Emily',  'Rojas', 'emily.rojas@yourschool.edu', 'teacher']
  ];
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  sheet.getRange(1, 4).setNote(
    'Role must be exactly "admin" or "teacher" (lowercase).\n' +
    'Admin users can access Settings and all reports.\n' +
    'Teacher users can submit referrals and view student profiles.'
  );
  sheet.autoResizeColumns(1, headers.length);
}

// ── ParentContacts sheet ──────────────────────────────────────
// Layout: StudentID | ParentGUID | FirstName | LastName | ParentEmail | Phone
function _setupParentContacts(ss) {
  if (ss.getSheetByName(SS_PARENT)) return;
  var sheet = ss.insertSheet(SS_PARENT);

  var headers = [
    'StudentID', 'ParentGUID', 'FirstName', 'LastName', 'ParentEmail', 'Phone'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  // S005 and S006 share a ParentGUID to demonstrate sibling-linking
  var data = [
    ['S001', 'pg-a1b2c3d4', 'John',  'Smith',    'john.smith@email.com',    '555-0101'],
    ['S002', 'pg-e5f6g7h8', 'Mary',  'Jones',    'mary.jones@email.com',    '555-0102'],
    ['S003', 'pg-i9j0k1l2', 'Robert','Taylor',   'robert.taylor@email.com', '555-0103'],
    ['S004', 'pg-m3n4o5p6', 'Sue',   'Anderson', 'sue.anderson@email.com',  '555-0104'],
    ['S005', 'pg-q7r8s9t0', 'Karen', 'Brown',    'karen.brown@email.com',   '555-0105'],
    ['S006', 'pg-q7r8s9t0', 'Karen', 'Brown',    'karen.brown@email.com',   '555-0105']
  ];
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  sheet.getRange(1, 2).setNote(
    'ParentGUID links siblings to the same parent.\n' +
    'Students sharing a ParentGUID will be shown together\n' +
    'when editing contact info in the web app.\n' +
    'Do not edit GUIDs manually — use the web app.'
  );
  sheet.autoResizeColumns(1, headers.length);
}

// =============================================================
// TRIGGER MANAGEMENT
// =============================================================

function setupDailyTrigger() {
  removeDailyTrigger();
  var cfg     = null;
  try { cfg = getConfig(); } catch(e) {}
  var timeStr = (cfg && cfg.emailSendTime) ? cfg.emailSendTime : '15:30';
  var hour    = parseInt(timeStr.split(':')[0], 10) || 15;
  ScriptApp.newTrigger('sendDailyParentEmails')
    .timeBased().everyDays(1).atHour(hour).create();
  SpreadsheetApp.getUi().alert(
    'Daily trigger created — parent emails will send at ' + timeStr + ' each day.'
  );
}

function removeDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyParentEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// =============================================================
// SEMESTER RESET (menu entry point)
// =============================================================

function resetSemesterPointsMenu() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Reset Semester Points',
    'This will reset ALL student point balances to the current ' +
    'Semester Start Points value.\n\nReferral history will NOT be deleted.' +
    '\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  var result = resetSemesterPoints();
  if (result.success) {
    ui.alert('Done', 'Reset ' + result.count + ' student balances.', ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.error || 'Unknown error.', ui.ButtonSet.OK);
  }
}