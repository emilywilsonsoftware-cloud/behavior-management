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
var SS_REFERRALS    = 'Referrals';
var SS_STUDENTS     = 'Students';
var SS_STAFF        = 'Staff';
var SS_PARENT       = 'ParentContacts';
var SS_CONFIG       = 'Config';
var SS_INFRACTIONS  = 'Infractions';
var SS_DELETION_LOG = 'DeletionLog';

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
    .addItem('Clear All Staff Role Caches', 'clearAllUserCachesMenu')
    .addSeparator()
    .addItem('Reset Semester Points (NEW SEMESTER)', 'resetSemesterPointsMenu')
    .addToUi();
}

/**
 * Spreadsheet-menu wrapper for clearAllUserCaches() (defined in Code.gs).
 * Use this any time you change a Role directly in the Staff sheet rather
 * than through Settings > Staff in the web app — that path already
 * clears the relevant cache automatically; editing the sheet directly
 * does not, so the old role can otherwise persist for up to 30 minutes.
 */
function clearAllUserCachesMenu() {
  var ui = SpreadsheetApp.getUi();
  clearAllUserCaches(); // defined in Code.gs — logs the count via Logger.log
  ui.alert(
    'Done',
    'Cleared cached roles for all staff. Everyone will be re-checked ' +
    'against the Staff sheet on their next page load.',
    ui.ButtonSet.OK
  );
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
  _setupDeletionLog(ss);

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
  // PointTierThresholds / PointTierColors seed the default 3-tier
  // green/amber/red point-balance coloring (70/40/0). Both lists must
  // stay the same length and be edited together via the Settings page —
  // editing them directly in the sheet risks producing mismatched rows
  // that getConfig() will reject and silently replace with this default.
  var headers = [
    'SchoolName', 'SemesterStartPoints', 'EmailNotificationsEnabled',
    'TeacherEmailNotificationsEnabled',
    'DailyEmailSendTime', 'EmailFooterText',
    'Locations', 'ContactTypes', 'NineWeeksStartDates',
    'PointTierThresholds', 'PointTierColors', 'PositiveCapPerNineWeeks'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var row2 = [
    'Our School', '100', 'Yes', 'Yes', '15:30',
    'If you have questions, please contact the school office.\n\n' +
    'This is an automated message — please do not reply.',
    'Classroom', 'Parent/Guardian', '',
    '70', 'green', '15'
  ];
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);

  var extraLocations     = ['Hallway', 'Cafeteria', 'Gym', 'Bus', 'Restroom', 'Office'];
  var extraContactTypes  = ['Administrator', 'Counselor', 'Case Manager'];
  var extraNineWeeks     = ['2024-08-01', '2024-10-15', '2025-01-08', '2025-03-18'];
  var extraThresholds    = ['40', '0'];
  var extraColors        = ['amber', 'red'];
  var maxExtra = Math.max(
    extraLocations.length, extraContactTypes.length, extraNineWeeks.length,
    extraThresholds.length, extraColors.length
  );
  for (var r = 0; r < maxExtra; r++) {
    sheet.getRange(r + 3, 1, 1, 11).setValues([[
      '', '', '', '', '', '',
      extraLocations[r]    || '',
      extraContactTypes[r] || '',
      extraNineWeeks[r]    || '',
      extraThresholds[r]   || '',
      extraColors[r]       || ''
    ]]);
  }
  sheet.setColumnWidth(6, 300);
  sheet.autoResizeColumns(1, 5);

  sheet.getRange(1, 10).setNote(
    'Pairs row-by-row with PointTierColors.\n' +
    'Defines the percentage thresholds for point-balance color tiers.\n' +
    'Edit via Settings > General in the web app, not directly here.'
  );
  sheet.getRange(1, 11).setNote(
    'Pairs row-by-row with PointTierThresholds.\n' +
    'Allowed values: green, blue, purple, amber, orange, red.\n' +
    'Edit via Settings > General in the web app, not directly here.'
  );
  sheet.getRange(1, 12).setNote(
    'Maximum total positive points (Write Off, Saturday School, etc.) ' +
    'a single student can be awarded per term, on the ' +
    'admin-only Positive Note page. Admins can still submit ' +
    'over this cap — it warns rather than blocks — but every override ' +
    'is noted on the referral for the record.'
  );
}

// ── Infractions sheet ─────────────────────────────────────────
function _setupInfractions(ss) {
  if (ss.getSheetByName(SS_INFRACTIONS)) return;
  var sheet = ss.insertSheet(SS_INFRACTIONS);

  var headers = ['InfractionName', 'PointValue', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  var data = [
    ['Dress Code Violation',         -3,  ''],
    ['Failure to Follow Directions', -5,  ''],
    ['Disruptive Behavior',          -5,  ''],
    ['Disrespect',                   -7,  ''],
    ['Bullying',                     -15, ''],
    ['Fighting',                     -20, ''],
    ['Vandalism',                    -15, ''],
    ['Harassment',                   -15, ''],
    ['Technology Violation',         -5,  ''],
    ['Tardy',                        -2,  ''],
    // Positive types — entered only via the admin-only "Award Positive
    // Points" tab (see getFormBootstrap in Code.gs), not the main
    // incident form. Add more rows here any time; no code changes
    // needed — the tab reads this list dynamically, same as the main
    // form already does for demerit types.
    ['Write Off',                     5,  ''],
    ['Saturday School',              10,  '']
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
    'PointValue', 'PointsBeforeReferral', 'PointsAfterReferral',
    'Description', 'IncludeDescriptionInEmail',
    'TeacherName', 'TeacherEmail',
    'ParentNotified', 'TeacherNotified', 'AdminNotes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, headers.length);
}

// ── Deletion log ──────────────────────────────────────────────
// Audit trail written by deleteReferral() in Code.gs — records who
// deleted a referral, when, and why, without keeping the deleted
// referral's full original content (location, description, etc.)
// lingering anywhere.
function _setupDeletionLog(ss) {
  if (ss.getSheetByName(SS_DELETION_LOG)) return;
  var sheet = ss.insertSheet(SS_DELETION_LOG);
  var headers = [
    'Timestamp', 'DeletedByName', 'DeletedByEmail',
    'ReferralID', 'StudentID', 'StudentName',
    'InfractionType', 'PointValue', 'IncidentDate', 'IncidentTime', 'Reason'
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
// Layout: StudentID | ContactGUID | Role | FirstName | LastName | Email
// A student can have MULTIPLE contact rows — Parent/Guardian,
// Administrator, Counselor, Case Manager — all of whom receive the
// same end-of-day referral email digest for that student. This list
// exists ONLY to drive referral email delivery; it is not synced from
// Infinite Campus or any other system and must be updated manually by
// an admin whenever who should receive referral emails changes.
function _setupParentContacts(ss) {
  if (ss.getSheetByName(SS_PARENT)) return;
  var sheet = ss.insertSheet(SS_PARENT);

  var headers = [
    'StudentID', 'ContactGUID', 'Role', 'FirstName', 'LastName', 'Email'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

  // S005 and S006 share a ContactGUID to demonstrate the "update for
  // siblings" feature — same parent, two kids at the school.
  var data = [
    ['S001', 'pg-a1b2c3d4', 'Parent/Guardian', 'John',   'Smith',    'john.smith@email.com'],
    ['S002', 'pg-e5f6g7h8', 'Parent/Guardian', 'Mary',   'Jones',    'mary.jones@email.com'],
    ['S003', 'pg-i9j0k1l2', 'Parent/Guardian', 'Robert', 'Taylor',   'robert.taylor@email.com'],
    ['S004', 'pg-m3n4o5p6', 'Parent/Guardian', 'Sue',    'Anderson', 'sue.anderson@email.com'],
    ['S005', 'pg-q7r8s9t0', 'Parent/Guardian', 'Karen',  'Brown',    'karen.brown@email.com'],
    ['S006', 'pg-q7r8s9t0', 'Parent/Guardian', 'Karen',  'Brown',    'karen.brown@email.com']
  ];
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  sheet.getRange(1, 2).setNote(
    'ContactGUID links rows that represent the SAME PERSON across\n' +
    'multiple students (e.g. one parent with two kids at the school).\n' +
    'Unrelated to Role. Do not edit GUIDs manually — use the web app.'
  );
  sheet.getRange(1, 3).setNote(
    'Must be exactly one of: Parent/Guardian, Administrator,\n' +
    'Counselor, Case Manager. Edit via the Student page in the web app.'
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