// ============================================================
// Behavior Tracker — Google Apps Script Web App
// Code.gs  v12 — Configurable point-balance color tiers
//
// CHANGES FROM v11:
//   • New Config columns: PointTierThresholds, PointTierColors —
//     parallel lists defining how many tiers exist, where each tier
//     starts (percentage of semester points), and which color from the
//     constrained palette each tier uses.
//   • POINT_COLOR_PALETTE constant — the fixed set of color keys/hex
//     values admins can choose from (green, blue, purple, amber, orange,
//     red). Prevents illegible or off-brand color choices.
//   • getConfig() now parses, validates, and returns pointTiers — falls
//     back to the original 3-tier default (green/amber/red at 70/40/0)
//     if the columns are missing, malformed, have fewer than 2 or more
//     than 6 tiers, or reference an unknown color key.
//   • getPointColor(pct, tiers) — shared helper resolving a percentage
//     to a hex color using a tier list; used wherever points are colored.
//   • getFormBootstrap(), getDashboardData(), getStudentProfile(),
//     getAllStudents(), getReportData() all now return pointTiers so
//     each page can resolve colors client-side with the same logic.
//   • getSettingsBootstrap() exposes the raw tier columns for editing.
//   • saveConfigSection() needed no changes — it already creates/updates
//     arbitrary Config columns generically based on payload keys.
// ============================================================

// ── SHEET NAMES ───────────────────────────────────────────────
var SHEET_REFERRALS   = 'Referrals';
var SHEET_STUDENTS    = 'Students';
var SHEET_STAFF       = 'Staff';
var SHEET_PARENT      = 'ParentContacts';
var SHEET_CONFIG      = 'Config';
var SHEET_INFRACTIONS = 'Infractions';

// ── CONFIG COLUMN NAMES ───────────────────────────────────────
var CONFIG_COL_LOCATIONS        = 'Locations';
// Redirections and Motivations intentionally removed — no longer used.
var CONFIG_COL_SCHOOL_NAME      = 'SchoolName';
var CONFIG_COL_SEMESTER_POINTS  = 'SemesterStartPoints';
var CONFIG_COL_EMAIL_ENABLED    = 'EmailNotificationsEnabled';
var CONFIG_COL_EMAIL_FOOTER     = 'EmailFooterText';
var CONFIG_COL_EMAIL_SEND_TIME  = 'DailyEmailSendTime';
var CONFIG_COL_NINE_WEEKS       = 'NineWeeksStartDates';
var CONFIG_COL_POINT_THRESHOLDS = 'PointTierThresholds';
var CONFIG_COL_POINT_COLORS     = 'PointTierColors';
// NOTE: CONFIG_COL_ADMIN_EMAILS intentionally removed.
// Admin role is now determined by Role = 'admin' in the Staff sheet.

// ── POINT BALANCE COLOR PALETTE (constrained — admins pick from this) ──
// Keeps color choices visually consistent with the rest of the app and
// prevents illegible or off-brand selections. Both the key (stored in
// Config) and hex value (used for rendering) live here as the single
// source of truth on the server; Settings.html mirrors this list for
// the picker UI.
var POINT_COLOR_PALETTE = {
  green:  '#10b981',
  blue:   '#3b82f6',
  purple: '#8b5cf6',
  amber:  '#eab308',
  orange: '#f97316',
  red:    '#ef4444'
};
var POINT_COLOR_DEFAULT = '#94a3b8'; // neutral gray fallback for unknown keys

// Default 3-tier setup, used whenever Config has no valid tier data yet.
var DEFAULT_POINT_TIERS = [
  { threshold: 70, color: 'green' },
  { threshold: 40, color: 'amber' },
  { threshold: 0,  color: 'red'   }
];

// ── STUDENTS SHEET COLUMN INDICES (0-based) ───────────────────
// Layout: StudentID | FirstName | LastName | MiddleName | Grade |
//         CurrentPoints | PointsLastUpdated
var STU_COL_ID          = 0;
var STU_COL_FIRST       = 1;
var STU_COL_LAST        = 2;
var STU_COL_MIDDLE      = 3;
var STU_COL_GRADE       = 4;
var STU_COL_POINTS      = 5;
var STU_COL_POINTS_DATE = 6;

// ── STAFF SHEET COLUMN INDICES (0-based) ─────────────────────
// Layout: FirstName | LastName | StaffEmail | Role
var STAFF_COL_FIRST = 0;
var STAFF_COL_LAST  = 1;
var STAFF_COL_EMAIL = 2;
var STAFF_COL_ROLE  = 3;

// ── PARENT CONTACTS SHEET COLUMN INDICES (0-based) ────────────
// Layout: StudentID | ParentGUID | FirstName | LastName | ParentEmail | Phone
var PARENT_COL_STUDENT_ID = 0;
var PARENT_COL_GUID       = 1;
var PARENT_COL_FIRST      = 2;
var PARENT_COL_LAST       = 3;
var PARENT_COL_EMAIL      = 4;
var PARENT_COL_PHONE      = 5;

// ── INFRACTIONS SHEET COLUMN NAMES ───────────────────────────
var INF_COL_NAME     = 'InfractionName';
var INF_COL_POINTS   = 'PointValue';
var INF_COL_SEVERITY = 'Severity';
var INF_COL_NOTES    = 'Notes';

// ── REFERRAL HEADERS ──────────────────────────────────────────
// ClassName, Redirections, PossibleMotivation intentionally removed — no longer used.
var REFERRAL_HEADERS = [
  'ID', 'Timestamp', 'StudentID', 'StudentName', 'Grade',
  'IncidentDate', 'IncidentTime', 'Location', 'InfractionType',
  'Severity', 'PointValue', 'PointsBeforeReferral', 'PointsAfterReferral',
  'Description', 'IncludeDescriptionInEmail',
  'TeacherName', 'TeacherEmail',
  'ParentNotified', 'TeacherNotified', 'Status', 'AdminNotes'
];

// ── EXECUTION CACHES (per-execution only — not persistent) ────
var _cfg  = null;
var _infs = null;
var _user = null;

// =============================================================
// ROUTING
// =============================================================

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'dashboard';
  var map  = {
    dashboard: 'Dashboard',
    form:      'Index',
    report:    'Report',
    student:   'Student',
    settings:  'Settings'
  };
  var file = map[page] || 'Dashboard';
  var base = ScriptApp.getService().getUrl();

  // ── Role gate for the referral form ──────────────────────────
  if (page === 'form') {
    var gateUser = getCurrentUser();
    if (gateUser.role !== 'admin' && gateUser.role !== 'teacher') {
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head>' +
        '<meta charset="UTF-8">' +
        '<meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<base target="_top">' +
        '<style>' +
        'body{margin:0;font-family:"Segoe UI",system-ui,sans-serif;background:#f1f5f9;}' +
        '.hdr{background:#1e3a5f;color:white;padding:0 20px;height:54px;' +
             'display:flex;align-items:center;font-weight:700;font-size:1.05rem;}' +
        '.wrap{max-width:500px;margin:60px auto;padding:0 16px;}' +
        '.card{background:white;border-radius:10px;padding:36px 32px;' +
              'box-shadow:0 1px 4px rgba(0,0,0,0.1);text-align:center;}' +
        '.icon{font-size:2.5rem;margin-bottom:14px;}' +
        'h2{color:#dc2626;font-size:1.1rem;margin-bottom:10px;}' +
        'p{color:#64748b;font-size:0.88rem;line-height:1.6;margin-bottom:20px;}' +
        'a{display:inline-block;padding:9px 22px;background:#1e3a5f;color:white;' +
          'border-radius:7px;text-decoration:none;font-size:0.88rem;font-weight:600;}' +
        '</style></head><body>' +
        '<div class="hdr">Behavior Tracker</div>' +
        '<div class="wrap"><div class="card">' +
        '<div class="icon">🔒</div>' +
        '<h2>Access Denied</h2>' +
        '<p>The referral form is only available to registered teachers and administrators.<br>' +
        'If you believe this is an error, please contact your school administrator.</p>' +
        '<a href="' + base + '?page=dashboard">Go to Dashboard</a>' +
        '</div></div></body></html>'
      ).setTitle('Access Denied')
       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  var pages = ['dashboard', 'form', 'student', 'report', 'settings'];
  var navHtml = HtmlService.createHtmlOutputFromFile('Nav').getContent();
  navHtml = navHtml.replace(/\{\{BASE\}\}/g, base);
  for (var i = 0; i < pages.length; i++) {
    var pg = pages[i];
    navHtml = navHtml.replace(
      '{{ACT_' + pg + '}}',
      pg === page ? 'active' : ''
    );
  }

  var tmpl = HtmlService.createTemplateFromFile(file);
  tmpl.navHtml = navHtml;

  return tmpl.evaluate()
    .setTitle('Behavior Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptURL(qs) {
  var url = ScriptApp.getService().getUrl();
  return qs ? url + '?' + qs : url;
}

// =============================================================
// UTILITIES — DATE FORMATTING
// =============================================================

function formatDateStr(val) {
  if (!val) return '';
  var d;
  if (val instanceof Date) {
    d = val;
  } else {
    d = new Date(val);
    if (isNaN(d.getTime())) return val.toString().trim();
  }
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
}

/**
 * Formats a time value read from a sheet cell into "HH:MM" (24-hr).
 * Time-only cells written from an <input type="time"> often get
 * auto-converted by Google Sheets into a Date object internally
 * (Sheets' time epoch, e.g. Dec 30 1899) even though they display
 * as a plain time in the UI. This handles both cases:
 *   - actual Date objects (read the hours/minutes directly)
 *   - plain strings already in "HH:MM" or "H:MM" format (returned as-is)
 *   - anything else falls back to the raw string, trimmed
 */
function formatTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return pad2(val.getHours()) + ':' + pad2(val.getMinutes());
  }
  var str = val.toString().trim();
  // Already a clean "H:MM" or "HH:MM" string — return as-is.
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(str)) {
    return str.slice(0, 5); // strip seconds if present, normalize to HH:MM
  }
  // Fallback — try parsing as a date in case it's an ISO-ish timestamp string.
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return pad2(parsed.getHours()) + ':' + pad2(parsed.getMinutes());
  }
  return str;
}

function formatDate(d) {
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
}

function pad2(n) { return n < 10 ? '0' + n : '' + n; }

function subtractDay(dateStr) {
  var d = new Date(dateStr + 'T12:00:00');
  d.setDate(d.getDate() - 1);
  return formatDate(d);
}

/**
 * Combines first and last name into "First Last" display format.
 * Trims whitespace and collapses any double spaces.
 */
function displayName(first, last) {
  return (((first || '').trim()) + ' ' + ((last || '').trim())).trim().replace(/\s+/g, ' ');
}

/**
 * Maps the stored severity value to the word shown on teacher-facing
 * surfaces (emails, etc). Major and Minor both collapse into the single
 * word "Negative" — teachers no longer see the Major/Minor distinction
 * anywhere. Positive stays Positive. An infraction with no severity
 * assigned (empty string) maps to an empty string here too, so it's
 * simply omitted from any "(Severity)" parenthetical in generated text.
 * The underlying stored value in the Referrals/Infractions sheets is
 * NEVER changed by this — it only affects what's displayed/sent.
 */
function displaySeverity(severity) {
  var s = (severity || '').toString().trim();
  if (s === 'Major' || s === 'Minor') return 'Negative';
  if (s === 'Positive') return 'Positive';
  return '';
}

// =============================================================
// CONFIG & INFRACTIONS READERS  (cached per execution)
// =============================================================

function getConfig() {
  if (_cfg) return _cfg;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) throw new Error('Config sheet not found. Please run Initial Setup.');

  var data = sheet.getDataRange().getValues();
  if (data.length < 1) throw new Error('Config sheet is empty. Please run Initial Setup.');

  var headers = data[0];
  var colMap  = {};

  headers.forEach(function(h, ci) {
    if (!h) return;
    var key = h.toString().trim();
    colMap[key] = [];
    for (var ri = 1; ri < data.length; ri++) {
      var v = data[ri][ci];
      if (v !== null && v !== undefined && v.toString().trim() !== '') {
        colMap[key].push(v.toString().trim());
      }
    }
  });

  function col(n)      { return colMap[n] || []; }
  function val(n, def) { var c = col(n); return c.length ? c[0] : def; }

  var nwStarts = col(CONFIG_COL_NINE_WEEKS)
    .map(function(d) { return d.trim(); })
    .filter(function(d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); })
    .sort();

  var nineWeeks = nwStarts.map(function(start, idx) {
    var endDate = nwStarts[idx + 1]
      ? subtractDay(nwStarts[idx + 1])
      : '9999-12-31';
    return { start: start, end: endDate, label: 'Nine Weeks ' + (idx + 1) };
  });

  var today = formatDate(new Date());
  var curNW = null;
  for (var i = 0; i < nineWeeks.length; i++) {
    if (today >= nineWeeks[i].start && today <= nineWeeks[i].end) {
      curNW = nineWeeks[i];
      break;
    }
  }

  var pointTiers = parsePointTiers(
    col(CONFIG_COL_POINT_THRESHOLDS),
    col(CONFIG_COL_POINT_COLORS)
  );

  _cfg = {
    locations:        col(CONFIG_COL_LOCATIONS),
    schoolName:       val(CONFIG_COL_SCHOOL_NAME,     'Our School'),
    semesterPoints:   parseInt(val(CONFIG_COL_SEMESTER_POINTS, '100'), 10) || 100,
    emailEnabled:     val(CONFIG_COL_EMAIL_ENABLED, 'Yes').toLowerCase() === 'yes',
    emailFooter:      col(CONFIG_COL_EMAIL_FOOTER).join('\n') ||
                        'If you have questions, please contact the school office.\n\n' +
                        'This is an automated message — please do not reply.',
    emailSendTime:    val(CONFIG_COL_EMAIL_SEND_TIME, '15:30'),
    nineWeeks:        nineWeeks,
    currentNineWeeks: curNW,
    pointTiers:       pointTiers,
    _raw:             colMap
    // NOTE: adminEmails intentionally absent — use getCurrentUser() for role checks.
  };
  return _cfg;
}

/**
 * Validates and normalizes raw Config column data into a clean tier list.
 * Falls back to DEFAULT_POINT_TIERS if the data is missing, malformed,
 * has fewer than 2 or more than 6 tiers, contains an unknown color key,
 * or any threshold isn't a number between 0 and 100.
 * Returned tiers are always sorted descending by threshold, and the
 * lowest tier's threshold is forced to 0 regardless of what was stored,
 * since the bottom tier must always catch every remaining percentage.
 */
function parsePointTiers(rawThresholds, rawColors) {
  try {
    if (!rawThresholds || !rawColors) return DEFAULT_POINT_TIERS;
    if (rawThresholds.length !== rawColors.length) return DEFAULT_POINT_TIERS;
    if (rawThresholds.length < 2 || rawThresholds.length > 6) return DEFAULT_POINT_TIERS;

    var tiers = [];
    for (var i = 0; i < rawThresholds.length; i++) {
      var t = parseFloat(rawThresholds[i]);
      var c = (rawColors[i] || '').toString().trim().toLowerCase();
      if (isNaN(t) || t < 0 || t > 100)        return DEFAULT_POINT_TIERS;
      if (!POINT_COLOR_PALETTE.hasOwnProperty(c)) return DEFAULT_POINT_TIERS;
      tiers.push({ threshold: t, color: c });
    }

    tiers.sort(function(a, b) { return b.threshold - a.threshold; });
    tiers[tiers.length - 1].threshold = 0; // bottom tier always catches the rest

    return tiers;
  } catch (e) {
    return DEFAULT_POINT_TIERS;
  }
}

/**
 * Resolves a percentage (0-100) to a hex color using a tier list.
 * Tiers must be sorted descending by threshold (parsePointTiers
 * guarantees this). Returns POINT_COLOR_DEFAULT if no tier matches,
 * which should only happen if tiers is empty.
 */
function getPointColor(pct, tiers) {
  var list = (tiers && tiers.length) ? tiers : DEFAULT_POINT_TIERS;
  for (var i = 0; i < list.length; i++) {
    if (pct >= list[i].threshold) {
      return POINT_COLOR_PALETTE[list[i].color] || POINT_COLOR_DEFAULT;
    }
  }
  return POINT_COLOR_DEFAULT;
}

function getInfractions() {
  if (_infs) return _infs;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_INFRACTIONS);
  if (!sheet) throw new Error('Infractions sheet not found. Please run Initial Setup.');

  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h) { return h.toString().trim(); });
  var ni   = hdrs.indexOf(INF_COL_NAME);
  var pi   = hdrs.indexOf(INF_COL_POINTS);
  var si   = hdrs.indexOf(INF_COL_SEVERITY);

  if (ni < 0) throw new Error('Infractions sheet missing InfractionName column.');

  _infs = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][ni] ? data[i][ni].toString().trim() : '';
    if (!name) continue;
    var pts = parseInt(data[i][pi], 10);
    _infs.push({
      name:       name,
      pointValue: isNaN(pts) ? 0 : pts,
      severity:   si >= 0 ? data[i][si].toString().trim() : ''
    });
  }
  _infs.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return _infs;
}

function getPointValue(infractionName) {
  var list = getInfractions();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === infractionName) return list[i].pointValue;
  }
  return 0;
}

// =============================================================
// IDENTITY & ROLES
// =============================================================

/**
 * Derives role by reading the Staff sheet directly.
 * Role column (STAFF_COL_ROLE) must contain 'admin' or 'teacher' (case-insensitive).
 * No longer references AdminEmails from Config.
 *
 * Two-level cache:
 *   Level 1 — _user var: free within a single GAS execution.
 *   Level 2 — CacheService.getUserCache() (30 min): persists across page
 *     loads. Invalidated by saveStaff() when the Staff sheet is modified.
 */
function getCurrentUser() {
  if (_user) return _user;

  var email    = Session.getActiveUser().getEmail();
  var cacheKey = 'bt_user_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

  try {
    var cached = CacheService.getUserCache().get(cacheKey);
    if (cached) {
      _user = JSON.parse(cached);
      return _user;
    }
  } catch (e) {
    Logger.log('CacheService read failed: ' + e.message);
  }

  // Cache miss — read Staff sheet once.
  var cfg   = getConfig();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var staff = ss.getSheetByName(SHEET_STAFF).getDataRange().getValues();

  var first = '';
  var last  = '';
  var role  = 'viewer';

  for (var i = 1; i < staff.length; i++) {
    var rowEmail = staff[i][STAFF_COL_EMAIL] ? staff[i][STAFF_COL_EMAIL].toString().trim() : '';
    if (rowEmail.toLowerCase() === email.toLowerCase()) {
      first = staff[i][STAFF_COL_FIRST] ? staff[i][STAFF_COL_FIRST].toString().trim() : '';
      last  = staff[i][STAFF_COL_LAST]  ? staff[i][STAFF_COL_LAST].toString().trim()  : '';
      var rowRole = staff[i][STAFF_COL_ROLE] ? staff[i][STAFF_COL_ROLE].toString().trim().toLowerCase() : '';
      if (rowRole === 'admin' || rowRole === 'teacher') {
        role = rowRole;
      }
      break;
    }
  }

  _user = {
    email:      email,
    name:       displayName(first, last) || email.split('@')[0],
    role:       role,
    isAdmin:    role === 'admin',
    isTeacher:  role === 'admin' || role === 'teacher',
    schoolName: cfg.schoolName
  };

  try {
    CacheService.getUserCache().put(cacheKey, JSON.stringify(_user), 1800);
  } catch (e) {
    Logger.log('CacheService write failed: ' + e.message);
  }

  return _user;
}

function invalidateUserCache(email) {
  try {
    var cacheKey = 'bt_user_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');
    CacheService.getUserCache().remove(cacheKey);
  } catch (e) {
    Logger.log('Cache invalidation failed for ' + email + ': ' + e.message);
  }
}

// =============================================================
// SETTINGS PAGE  (admin only)
// =============================================================

function getSettingsBootstrap() {
  var cfg  = getConfig();
  var user = getCurrentUser();
  var raw  = cfg._raw || {};

  return {
    user: user,
    colorPalette: POINT_COLOR_PALETTE, // {key: hex} — drives the Settings color picker
    config: {
      SchoolName:                raw[CONFIG_COL_SCHOOL_NAME]       || [],
      SemesterStartPoints:       raw[CONFIG_COL_SEMESTER_POINTS]   || [],
      EmailNotificationsEnabled: raw[CONFIG_COL_EMAIL_ENABLED]     || [],
      DailyEmailSendTime:        raw[CONFIG_COL_EMAIL_SEND_TIME]   || [],
      EmailFooterText:           raw[CONFIG_COL_EMAIL_FOOTER]      || [],
      Locations:                 raw[CONFIG_COL_LOCATIONS]         || [],
      NineWeeksStartDates:       raw[CONFIG_COL_NINE_WEEKS]        || [],
      // Tier columns reflect the VALIDATED tiers, not raw sheet content —
      // so if the sheet had malformed data, the editor opens already
      // showing the safe 3-tier default rather than broken values.
      PointTierThresholds:       cfg.pointTiers.map(function(t) { return t.threshold; }),
      PointTierColors:           cfg.pointTiers.map(function(t) { return t.color; })
      // AdminEmails intentionally absent — managed via Staff manager
    }
  };
}

function saveConfigSection(payload) {
  var user = getCurrentUser();
  if (!user || user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  // ── Tier-specific validation ──────────────────────────────────
  // Catches bad data before it's written, so a malformed save fails
  // loudly here rather than silently falling back to defaults the
  // next time getConfig() reads it back.
  if (payload.hasOwnProperty(CONFIG_COL_POINT_THRESHOLDS) || payload.hasOwnProperty(CONFIG_COL_POINT_COLORS)) {
    var thresholds = payload[CONFIG_COL_POINT_THRESHOLDS] || [];
    var colors     = payload[CONFIG_COL_POINT_COLORS]     || [];

    if (thresholds.length !== colors.length) {
      return { success: false, error: 'Point tier thresholds and colors must have the same number of entries.' };
    }
    if (thresholds.length < 2) {
      return { success: false, error: 'At least 2 point tiers are required.' };
    }
    if (thresholds.length > 6) {
      return { success: false, error: 'No more than 6 point tiers are allowed.' };
    }
    for (var ti = 0; ti < thresholds.length; ti++) {
      var tVal = parseFloat(thresholds[ti]);
      if (isNaN(tVal) || tVal < 0 || tVal > 100) {
        return { success: false, error: 'Each tier threshold must be a number between 0 and 100.' };
      }
      var cVal = (colors[ti] || '').toString().trim().toLowerCase();
      if (!POINT_COLOR_PALETTE.hasOwnProperty(cVal)) {
        return { success: false, error: 'Unknown color "' + colors[ti] + '". Choose from: ' + Object.keys(POINT_COLOR_PALETTE).join(', ') + '.' };
      }
    }
  }

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CONFIG);
    if (!sheet) return { success: false, error: 'Config sheet not found.' };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var keys    = Object.keys(payload);

    for (var k = 0; k < keys.length; k++) {
      var key    = keys[k];
      var values = payload[key];

      var colNum = -1;
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].toString().trim() === key) { colNum = h + 1; break; }
      }
      if (colNum < 0) {
        colNum = lastCol + 1;
        sheet.getRange(1, colNum).setValue(key);
        lastCol++;
      }

      if (lastRow > 1) {
        sheet.getRange(2, colNum, lastRow - 1, 1).clearContent();
      }
      for (var v = 0; v < values.length; v++) {
        sheet.getRange(v + 2, colNum).setValue(values[v]);
      }
    }

    SpreadsheetApp.flush();
    _cfg  = null;
    _user = null;

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================
// SECURITY HELPERS
// =============================================================

function sanitizeText(val) {
  if (val === null || val === undefined) return '';
  return val.toString()
    .replace(/[<>]/g, '')
    .replace(/javascript:/gi, '')
    .trim()
    .slice(0, 1000);
}

// =============================================================
// FORM BOOTSTRAP
// =============================================================

function getFormBootstrap() {
  var cfg  = getConfig();
  var infs = getInfractions();
  var user = getCurrentUser();

  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied. You must be a registered teacher or admin to submit referrals.');
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var students = [];
  for (var i = 1; i < stuData.length; i++) {
    var pts = parseInt(stuData[i][STU_COL_POINTS], 10);
    if (isNaN(pts)) pts = cfg.semesterPoints;
    var pct = cfg.semesterPoints > 0 ? Math.round(pts / cfg.semesterPoints * 100) : 0;
    students.push({
      studentId:     stuData[i][STU_COL_ID].toString(),
      studentName:   displayName(stuData[i][STU_COL_FIRST], stuData[i][STU_COL_LAST]),
      grade:         stuData[i][STU_COL_GRADE].toString(),
      currentPoints: pts,
      pct:           pct,
      pointColor:    getPointColor(pct, cfg.pointTiers)
    });
  }
  students.sort(function(a, b) { return a.studentName.localeCompare(b.studentName); });

  return {
    user:             user,
    schoolName:       cfg.schoolName,
    semesterPoints:   cfg.semesterPoints,
    emailEnabled:     cfg.emailEnabled,
    locations:        cfg.locations,
    currentNineWeeks: cfg.currentNineWeeks,
    nineWeeks:        cfg.nineWeeks,
    pointTiers:       cfg.pointTiers,
    infractions:      infs.map(function(inf) {
      return { name: inf.name, pointValue: inf.pointValue, severity: inf.severity };
    }),
    students: students
  };
}

function getStudentFormCard(studentId) {
  var user = getCurrentUser();
  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied.');
  }

  var cfg     = getConfig();
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();

  var student = null;
  for (var i = 1; i < stuData.length; i++) {
    if (stuData[i][STU_COL_ID].toString() === studentId.toString()) {
      var pts = parseInt(stuData[i][STU_COL_POINTS], 10);
      student = {
        studentId:     stuData[i][STU_COL_ID].toString(),
        studentName:   displayName(stuData[i][STU_COL_FIRST], stuData[i][STU_COL_LAST]),
        grade:         stuData[i][STU_COL_GRADE].toString(),
        currentPoints: isNaN(pts) ? cfg.semesterPoints : pts
      };
      break;
    }
  }
  if (!student) return { error: 'Student not found: ' + studentId };

  var nw      = cfg.currentNineWeeks;
  var nwStart = nw ? nw.start : null;
  var nwEnd   = nw ? nw.end   : null;
  var nwLabel = nw ? nw.label : 'Current Nine Weeks';

  var hdrs = (refData && refData.length > 0) ? refData[0] : [];
  var ci   = {};
  for (var h = 0; h < hdrs.length; h++) {
    if (hdrs[h]) ci[hdrs[h].toString().trim()] = h;
  }

  var nwRefs = [];
  for (var r = 1; r < refData.length; r++) {
    var row = refData[r];
    if ((row[ci['StudentID']] || '').toString() !== studentId.toString()) continue;

    var incDate = formatDateStr(row[ci['IncidentDate']]);
    if (nwStart && incDate < nwStart) continue;
    if (nwEnd   && incDate > nwEnd)   continue;

    nwRefs.push({
      id:             row[ci['ID']] !== undefined ? row[ci['ID']].toString() : '',
      incidentDate:   incDate,
      incidentTime:   formatTimeStr(row[ci['IncidentTime']]),
      infractionType: row[ci['InfractionType']] ? row[ci['InfractionType']].toString() : '',
      severity:       row[ci['Severity']]       ? row[ci['Severity']].toString()       : '',
      pointValue:     parseFloat(row[ci['PointValue']]) || 0,
      teacherName:    row[ci['TeacherName']]    ? row[ci['TeacherName']].toString()    : '',
      status:         row[ci['Status']]         ? row[ci['Status']].toString()         : 'Open',
      description:    row[ci['Description']]    ? row[ci['Description']].toString()    : ''
    });
  }

  nwRefs.sort(function(a, b) {
    return b.incidentDate < a.incidentDate ? -1 : b.incidentDate > a.incidentDate ? 1 : 0;
  });

  return {
    student:        student,
    nwRefs:         nwRefs,
    nwLabel:        nwLabel,
    nwStart:        nwStart,
    nwEnd:          nwEnd,
    semesterPoints: cfg.semesterPoints
  };
}

// =============================================================
// SUBMIT REFERRALS
// =============================================================

function submitReferrals(referrals) {
  var user = getCurrentUser();
  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied. You must be a registered teacher or admin to submit referrals.');
  }

  if (!Array.isArray(referrals) || referrals.length === 0) {
    throw new Error('No referrals provided.');
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName(SHEET_REFERRALS);
  var stuSheet = ss.getSheetByName(SHEET_STUDENTS);
  var cfg      = getConfig();

  ensureHeaders(refSheet, REFERRAL_HEADERS);

  var stuData     = stuSheet.getDataRange().getValues();
  var validStuIds = {};
  for (var vi = 1; vi < stuData.length; vi++) {
    validStuIds[stuData[vi][STU_COL_ID].toString()] = true;
  }

  var REQUIRED_FIELDS = ['studentId', 'infractionType', 'location', 'incidentDate', 'incidentTime'];

  var lastId     = getLastId(refSheet);
  var timestamp  = new Date();
  var results    = { saved: 0, errors: [] };
  var emailQueue = [];

  for (var i = 0; i < referrals.length; i++) {
    try {
      var r = referrals[i];

      for (var fi = 0; fi < REQUIRED_FIELDS.length; fi++) {
        var field = REQUIRED_FIELDS[fi];
        if (!r[field] || r[field].toString().trim() === '') {
          throw new Error('Missing required field: ' + field);
        }
      }

      var sidStr = r.studentId.toString().trim();
      if (!validStuIds[sidStr]) {
        throw new Error('Student ID not found in roster: ' + sidStr);
      }

      var safeDescription = sanitizeText(r.description);
      var includeDescInEmail = r.includeDescriptionInEmail === true;

      var id = ++lastId;

      var pointValue   = getPointValue(r.infractionType);
      var stuIdx       = findStudentRow(stuData, sidStr);
      var pointsBefore = cfg.semesterPoints;

      if (stuIdx >= 0) {
        var ex = stuData[stuIdx][STU_COL_POINTS];
        if (ex !== '' && ex !== null) {
          var parsed = parseInt(ex, 10);
          if (!isNaN(parsed)) pointsBefore = parsed;
        }
      }

      var pointsAfter = Math.max(0, pointsBefore + pointValue);

      var newRowNum = refSheet.appendRow([
        id, timestamp,
        sidStr,
        sanitizeText(r.studentName),
        sanitizeText(r.grade),
        r.incidentDate, r.incidentTime,
        sanitizeText(r.location),
        sanitizeText(r.infractionType),
        sanitizeText(r.severity),
        pointValue, pointsBefore, pointsAfter,
        safeDescription,
        includeDescInEmail ? 'Yes' : 'No',
        sanitizeText(r.teacherName),
        user.email,
        'No', 'No', 'Open', ''
      ]).getLastRow();

      // Force the IncidentTime cell to plain-text format and re-write the
      // value as a literal string. Without this, Google Sheets auto-detects
      // the "HH:MM" string as a time-of-day value and silently converts it
      // to a Date/time serial — which then reads back incorrectly anywhere
      // the column is used (Report, Student profile, nine-weeks panel, etc).
      var timeCol = REFERRAL_HEADERS.indexOf('IncidentTime') + 1;
      var timeCell = refSheet.getRange(newRowNum, timeCol);
      timeCell.setNumberFormat('@STRING@');
      timeCell.setValue(r.incidentTime);

      if (stuIdx >= 0) {
        stuSheet.getRange(stuIdx + 1, STU_COL_POINTS + 1).setValue(pointsAfter);
        stuSheet.getRange(stuIdx + 1, STU_COL_POINTS_DATE + 1).setValue(timestamp);
        stuData[stuIdx][STU_COL_POINTS]      = pointsAfter;
        stuData[stuIdx][STU_COL_POINTS_DATE] = timestamp;
      }

      results.saved++;
      emailQueue.push({
        referral:     r,
        id:           id,
        pointValue:   pointValue,
        pointsBefore: pointsBefore,
        pointsAfter:  pointsAfter
      });

    } catch (err) {
      results.errors.push('Row ' + (i + 1) + ': ' + err.message);
    }
  }

  SpreadsheetApp.flush();

  for (var j = 0; j < emailQueue.length; j++) {
    var job = emailQueue[j];
    try {
      sendTeacherConfirmation(job.referral, job.id, job.pointValue,
                              job.pointsBefore, job.pointsAfter);
    } catch (emailErr) {
      Logger.log('Teacher email failed for #' + job.id + ': ' + emailErr.message);
    }
  }

  return results;
}

function findStudentRow(stuData, studentId) {
  for (var i = 1; i < stuData.length; i++) {
    if (stuData[i][STU_COL_ID].toString() === studentId.toString()) return i;
  }
  return -1;
}

// =============================================================
// TEACHER CONFIRMATION EMAIL
// =============================================================

function sendTeacherConfirmation(referral, referralId, pointValue, pointsBefore, pointsAfter) {
  try {
    if (!referral.teacherEmail) return;
    var cfg    = getConfig();
    var ptLine = pointValue < 0
      ? 'Points Deducted: ' + Math.abs(pointValue) + ' pts  (' + pointsBefore + ' → ' + pointsAfter + ')'
      : 'Points Added:    +' + pointValue + ' pts  (' + pointsBefore + ' → ' + pointsAfter + ')';

    var subject = '✓ Referral Saved — ' + referral.studentName + ' — #' + referralId;
    var dispSev = displaySeverity(referral.severity);
    var sevSuffix = dispSev ? ' (' + dispSev + ')' : '';
    var body    =
      'Hello ' + referral.teacherName + ',\n\n' +
      'Your referral has been saved.\n\n' +
      'Referral #:      ' + referralId + '\n' +
      'Student:         ' + referral.studentName + ' (Grade ' + referral.grade + ')\n' +
      'Date:            ' + referral.incidentDate + ' at ' + referral.incidentTime + '\n' +
      'Location:        ' + referral.location + '\n' +
      'Infraction:      ' + referral.infractionType + sevSuffix + '\n\n' +
      ptLine + '\n' +
      'Current Balance: ' + pointsAfter + ' pts\n\n' +
      (cfg.emailEnabled
        ? 'A parent notification will be sent at the end of the day.'
        : 'Parent notifications are currently disabled.') + '\n\n' +
      'This is an automated message from ' + cfg.schoolName + '. Do not reply.\n';

    GmailApp.sendEmail(referral.teacherEmail, subject, body);
    markColumn(referralId, 'TeacherNotified', 'Yes');
  } catch (err) {
    Logger.log('Teacher email failed for #' + referralId + ': ' + err.message);
  }
}

// =============================================================
// BATCHED END-OF-DAY PARENT EMAILS
// =============================================================

function sendDailyParentEmails() {
  var triggerEmail = Session.getActiveUser().getEmail();
  if (triggerEmail) {
    var callerUser = getCurrentUser();
    if (callerUser.role !== 'admin') {
      return { sent: 0, errors: ['Admin access required.'] };
    }
  }

  var cfg = getConfig();
  if (!cfg.emailEnabled) {
    Logger.log('Parent emails disabled in Config — skipping daily send.');
    return { sent: 0, errors: [] };
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName(SHEET_REFERRALS);
  var data     = refSheet.getDataRange().getValues();
  var today    = formatDate(new Date());

  var idIdx     = REFERRAL_HEADERS.indexOf('ID');
  var sidIdx    = REFERRAL_HEADERS.indexOf('StudentID');
  var snIdx     = REFERRAL_HEADERS.indexOf('StudentName');
  var grIdx     = REFERRAL_HEADERS.indexOf('Grade');
  var dtIdx     = REFERRAL_HEADERS.indexOf('IncidentDate');
  var tmIdx     = REFERRAL_HEADERS.indexOf('IncidentTime');
  var locIdx    = REFERRAL_HEADERS.indexOf('Location');
  var infIdx    = REFERRAL_HEADERS.indexOf('InfractionType');
  var sevIdx    = REFERRAL_HEADERS.indexOf('Severity');
  var ptIdx     = REFERRAL_HEADERS.indexOf('PointValue');
  var afIdx     = REFERRAL_HEADERS.indexOf('PointsAfterReferral');
  var tchIdx    = REFERRAL_HEADERS.indexOf('TeacherName');
  var descIdx   = REFERRAL_HEADERS.indexOf('Description');
  var incEmIdx  = REFERRAL_HEADERS.indexOf('IncludeDescriptionInEmail');
  var pnIdx     = REFERRAL_HEADERS.indexOf('ParentNotified');

  var contacts  = ss.getSheetByName(SHEET_PARENT).getDataRange().getValues();
  var parentMap = {};
  for (var p = 1; p < contacts.length; p++) {
    var sid = contacts[p][PARENT_COL_STUDENT_ID].toString().trim();
    if (sid && contacts[p][PARENT_COL_EMAIL]) {
      parentMap[sid] = {
        parentFirstName: contacts[p][PARENT_COL_FIRST] ? contacts[p][PARENT_COL_FIRST].toString().trim() : '',
        parentLastName:  contacts[p][PARENT_COL_LAST]  ? contacts[p][PARENT_COL_LAST].toString().trim()  : '',
        parentEmail:     contacts[p][PARENT_COL_EMAIL].toString().trim()
      };
    }
  }

  var grouped    = {};
  var rowIndices = {};

  for (var r = 1; r < data.length; r++) {
    var row      = data[r];
    var incDate  = formatDateStr(row[dtIdx]);
    var notified = row[pnIdx] ? row[pnIdx].toString() : '';

    if (incDate !== today) continue;
    if (notified === 'Yes') continue;

    var studentId = row[sidIdx] ? row[sidIdx].toString() : '';
    if (!studentId) continue;

    if (!grouped[studentId]) {
      grouped[studentId]    = [];
      rowIndices[studentId] = [];
    }

    var pts    = parseInt(row[ptIdx], 10) || 0;
    var ptsStr = pts > 0 ? '+' + pts : pts.toString();
    var includeDesc = (row[incEmIdx] || '').toString().trim().toLowerCase() === 'yes';

    grouped[studentId].push({
      id:             row[idIdx],
      studentName:    row[snIdx],
      grade:          row[grIdx],
      incidentDate:   incDate,
      incidentTime:   formatTimeStr(row[tmIdx]),
      location:       row[locIdx] ? row[locIdx].toString() : '',
      infractionType: row[infIdx] ? row[infIdx].toString() : '',
      severity:       row[sevIdx] ? row[sevIdx].toString() : '',
      pointValue:     pts,
      pointsStr:      ptsStr,
      pointsAfter:    row[afIdx] !== undefined ? row[afIdx] : '',
      teacherName:    row[tchIdx] ? row[tchIdx].toString() : '',
      // Description only carried through if the teacher opted in at submit time.
      description:    includeDesc && row[descIdx] ? row[descIdx].toString() : ''
    });
    rowIndices[studentId].push(r + 1);
  }

  var sent   = 0;
  var errors = [];

  Object.keys(grouped).forEach(function(studentId) {
    try {
      var contact = parentMap[studentId];
      if (!contact || !contact.parentEmail) return;

      var refs        = grouped[studentId];
      var studentName = refs[0].studentName;
      var grade       = refs[0].grade;
      var salutation  = contact.parentFirstName || displayName(contact.parentFirstName, contact.parentLastName) || 'Parent/Guardian';

      var subject = 'Daily Behavior Summary — ' + studentName + ' — ' + today;
      var body    =
        'Dear ' + salutation + ',\n\n' +
        'This is the daily behavior summary from ' + cfg.schoolName +
        ' for your student.\n\n' +
        'Student: ' + studentName + ' (Grade ' + grade + ')\n' +
        'Date:    ' + today + '\n\n';

      body += refs.length === 1
        ? '── Referral Received ────────────────────\n'
        : '── ' + refs.length + ' Referrals Received ─────────────────\n';

      refs.forEach(function(ref, idx) {
        if (refs.length > 1) body += '\nReferral ' + (idx + 1) + ':\n';
        var dispSev = displaySeverity(ref.severity);
        var sevSuffix = dispSev ? ' (' + dispSev + ')' : '';
        body +=
          'Time:        ' + ref.incidentTime + '\n' +
          'Location:    ' + ref.location + '\n' +
          'Infraction:  ' + ref.infractionType + sevSuffix + '\n' +
          'Teacher:     ' + ref.teacherName + '\n' +
          'Points:      ' + ref.pointsStr + '  (Balance: ' + ref.pointsAfter + ' pts)\n';
        if (ref.description) {
          body += 'Notes:       ' + ref.description + '\n';
        }
        body += '\n';
      });

      body += '─────────────────────────────────────────\n';
      body += 'Current Point Balance: ' + refs[refs.length - 1].pointsAfter + ' pts\n\n';
      body += cfg.emailFooter + '\n';

      GmailApp.sendEmail(contact.parentEmail, subject, body);
      sent++;

      rowIndices[studentId].forEach(function(rowNum) {
        refSheet.getRange(rowNum, pnIdx + 1).setValue('Yes');
      });

    } catch (err) {
      errors.push('StudentID ' + studentId + ': ' + err.message);
      Logger.log('Daily parent email error: ' + err.message);
    }
  });

  Logger.log('Daily parent emails: sent=' + sent + ', errors=' + errors.length);
  return { sent: sent, errors: errors };
}

// =============================================================
// MARK COLUMN UTILITY
// =============================================================

function markColumn(referralId, columnName, value) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REFERRALS);
  var data   = sheet.getDataRange().getValues();
  var idCol  = REFERRAL_HEADERS.indexOf('ID');
  var tgtCol = REFERRAL_HEADERS.indexOf(columnName);
  if (tgtCol < 0) return;
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == referralId) {
      sheet.getRange(i + 1, tgtCol + 1).setValue(value);
      return;
    }
  }
}

// =============================================================
// DASHBOARD DATA
// =============================================================

function getDashboardData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var cfg     = getConfig();
  var user    = getCurrentUser();

  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied.');
  }

  var now       = new Date();
  var today     = formatDate(now);
  var day7ago   = formatDate(new Date(now.getTime() - 7 * 86400000));
  var thisMonth = now.getMonth();
  var thisYear  = now.getFullYear();

  var hdrs = (refData && refData.length > 0) ? refData[0] : [];
  var ci   = {};
  for (var h = 0; h < hdrs.length; h++) {
    if (hdrs[h]) ci[hdrs[h].toString().trim()] = h;
  }

  var mMap   = {};
  var infMap = {};
  var locMap = {};
  var tchMap = {};

  var totalReferrals = 0;
  var openReferrals  = 0;
  var todayReferrals = 0;
  var weekReferrals  = 0;
  var negativeCount  = 0; // Major + Minor combined — teachers see one "Negative" bucket

  var allRefs = [];

  for (var r = 1; r < refData.length; r++) {
    var row  = refData[r];
    var inc  = formatDateStr(row[ci['IncidentDate']]);
    var sev  = row[ci['Severity']] ? row[ci['Severity']].toString() : '';
    var stat = row[ci['Status']]   ? row[ci['Status']].toString()   : '';
    var ts   = row[ci['Timestamp']];

    totalReferrals++;
    if (stat !== 'Resolved') openReferrals++;
    if (inc === today)       todayReferrals++;
    if (inc >= day7ago)      weekReferrals++;
    if (sev === 'Major' || sev === 'Minor') negativeCount++;

    var d = ts instanceof Date ? ts : (ts ? new Date(ts) : null);
    if (d && !isNaN(d.getTime())) {
      var mk = d.getFullYear() + '-' + pad2(d.getMonth() + 1);
      mMap[mk] = (mMap[mk] || 0) + 1;

      if (d.getMonth() === thisMonth && d.getFullYear() === thisYear) {
        var tch = row[ci['TeacherName']];
        if (tch) tchMap[tch] = (tchMap[tch] || 0) + 1;
      }
    }

    var inf = row[ci['InfractionType']];
    if (inf) infMap[inf] = (infMap[inf] || 0) + 1;

    var loc = row[ci['Location']];
    if (loc) locMap[loc] = (locMap[loc] || 0) + 1;

    allRefs.push({
      id:             row[ci['ID']] !== undefined ? row[ci['ID']].toString() : '',
      studentId:      row[ci['StudentID']]   ? row[ci['StudentID']].toString()   : '',
      studentName:    row[ci['StudentName']] ? row[ci['StudentName']].toString() : '',
      grade:          row[ci['Grade']]       ? row[ci['Grade']].toString()       : '',
      infractionType: inf ? inf.toString() : '',
      severity:       sev,
      pointValue:     parseFloat(row[ci['PointValue']]) || 0,
      incidentDate:   inc,
      incidentTime:   formatTimeStr(row[ci['IncidentTime']]),
      teacherName:    row[ci['TeacherName']] ? row[ci['TeacherName']].toString() : '',
      status:         stat
    });
  }

  var monthlyTrend = [];
  for (var m = 11; m >= 0; m--) {
    var md  = new Date(now.getFullYear(), now.getMonth() - m, 1);
    var mk2 = md.getFullYear() + '-' + pad2(md.getMonth() + 1);
    var ml  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][md.getMonth()] +
              ' ' + md.getFullYear().toString().slice(2);
    monthlyTrend.push({ label: ml, count: mMap[mk2] || 0 });
  }

  var infractionBreakdown = Object.keys(infMap)
    .map(function(k) { return { name: k, count: infMap[k] }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 8);

  var locationBreakdown = Object.keys(locMap)
    .map(function(k) { return { name: k, count: locMap[k] }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 6);

  var teacherActivity = Object.keys(tchMap)
    .map(function(k) { return { name: k, count: tchMap[k] }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 8);

  var recentReferrals = allRefs
    .filter(function(r) { return r.incidentDate; })
    .sort(function(a, b) {
      return a.incidentDate > b.incidentDate ? -1 :
             a.incidentDate < b.incidentDate ?  1 : 0;
    })
    .slice(0, 10);

  var sp     = cfg.semesterPoints;
  var atRisk = [];
  for (var s = 1; s < stuData.length; s++) {
    var pts = parseInt(stuData[s][STU_COL_POINTS], 10);
    if (isNaN(pts)) pts = sp;
    var pct = sp > 0 ? Math.round(pts / sp * 100) : 0;
    atRisk.push({
      studentId:     stuData[s][STU_COL_ID].toString(),
      studentName:   displayName(stuData[s][STU_COL_FIRST], stuData[s][STU_COL_LAST]),
      grade:         stuData[s][STU_COL_GRADE].toString(),
      currentPoints: pts,
      pct:           pct,
      pointColor:    getPointColor(pct, cfg.pointTiers)
    });
  }
  atRisk.sort(function(a, b) { return a.currentPoints - b.currentPoints; });

  // "At risk" means falling in the lowest configured tier — uses
  // whatever threshold the admin has set, not a hardcoded percentage.
  var lowestTierThreshold = cfg.pointTiers[cfg.pointTiers.length - 1].threshold;
  // The lowest tier's threshold is always 0, so "at risk" actually means
  // falling below the SECOND-lowest tier's threshold (i.e. not yet in
  // a "good" or "middle" tier). Guard for the 2-tier minimum case.
  var atRiskThreshold = cfg.pointTiers.length > 1
    ? cfg.pointTiers[cfg.pointTiers.length - 2].threshold
    : 100;
  var atRiskCount = 0;
  for (var ar = 0; ar < atRisk.length; ar++) {
    if (atRisk[ar].pct < atRiskThreshold) atRiskCount++;
  }
  atRisk = atRisk.slice(0, 10);

  return {
    user:                user,
    schoolName:          cfg.schoolName,
    semesterPoints:      sp,
    currentNineWeeks:    cfg.currentNineWeeks,
    pointTiers:          cfg.pointTiers,
    stats: {
      totalReferrals:  totalReferrals,
      openReferrals:   openReferrals,
      todayReferrals:  todayReferrals,
      weekReferrals:   weekReferrals,
      negativeCount:   negativeCount,
      totalStudents:   stuData.length - 1,
      atRiskCount:     atRiskCount
    },
    recentReferrals:     recentReferrals,
    atRisk:              atRisk,
    monthlyTrend:        monthlyTrend,
    infractionBreakdown: infractionBreakdown,
    locationBreakdown:   locationBreakdown,
    teacherActivity:     teacherActivity
  };
}

// =============================================================
// STUDENT PROFILE DATA
// =============================================================

function getStudentProfile(studentId) {
  var user = getCurrentUser();
  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied.');
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
  var pData   = ss.getSheetByName(SHEET_PARENT).getDataRange().getValues();
  var cfg     = getConfig();

  var hdrs = (refData && refData.length > 0) ? refData[0] : [];
  var ci   = {};
  for (var h = 0; h < hdrs.length; h++) {
    if (hdrs[h]) ci[hdrs[h].toString().trim()] = h;
  }

  if (ci['StudentID'] === undefined) {
    return {
      student:             null,
      referrals:           [],
      pointTimeline:       [],
      infractionBreakdown: [],
      nwSummary:           [],
      nineWeeks:           cfg.nineWeeks,
      currentNineWeeks:    cfg.currentNineWeeks,
      semesterPoints:      cfg.semesterPoints,
      schoolName:          cfg.schoolName,
      pointTiers:          cfg.pointTiers,
      summary: { total: 0, open: 0, major: 0, minor: 0, positive: 0 }
    };
  }

  var student = null;
  for (var s = 1; s < stuData.length; s++) {
    if (stuData[s][STU_COL_ID].toString() === studentId.toString()) {
      var pts = parseInt(stuData[s][STU_COL_POINTS], 10);
      var curPts = isNaN(pts) ? cfg.semesterPoints : pts;
      var curPct = cfg.semesterPoints > 0 ? Math.round(curPts / cfg.semesterPoints * 100) : 0;
      student = {
        studentId:         stuData[s][STU_COL_ID].toString(),
        studentName:   displayName(stuData[s][STU_COL_FIRST], stuData[s][STU_COL_LAST]),
        grade:             stuData[s][STU_COL_GRADE].toString(),
        currentPoints:     curPts,
        pct:               curPct,
        pointColor:        getPointColor(curPct, cfg.pointTiers),
        pointsLastUpdated: formatDateStr(stuData[s][STU_COL_POINTS_DATE])
      };
      break;
    }
  }
  if (!student) return { error: 'Student ' + studentId + ' not found.' };

  for (var p = 1; p < pData.length; p++) {
    if (pData[p][PARENT_COL_STUDENT_ID].toString() === studentId.toString()) {
      student.parentGuid  = pData[p][PARENT_COL_GUID]  ? pData[p][PARENT_COL_GUID].toString()  : '';
      student.parentFirst = pData[p][PARENT_COL_FIRST] ? pData[p][PARENT_COL_FIRST].toString() : '';
      student.parentLast  = pData[p][PARENT_COL_LAST]  ? pData[p][PARENT_COL_LAST].toString()  : '';
      student.parentName  = displayName(student.parentFirst, student.parentLast);
      student.parentEmail = pData[p][PARENT_COL_EMAIL] ? pData[p][PARENT_COL_EMAIL].toString() : '';
      student.parentPhone = pData[p][PARENT_COL_PHONE] ? pData[p][PARENT_COL_PHONE].toString() : '';
      break;
    }
  }

  function str(v) { return (v === null || v === undefined) ? '' : v.toString(); }
  function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

  var referrals = [];
  for (var r = 1; r < refData.length; r++) {
    var row = refData[r];
    if ((row[ci['StudentID']] || '').toString() !== studentId.toString()) continue;

    var inc = formatDateStr(row[ci['IncidentDate']]);
    var ts  = row[ci['Timestamp']];

    referrals.push({
      ID:                   str(row[ci['ID']]),
      IncidentDate:         inc,
      IncidentTime:         formatTimeStr(row[ci['IncidentTime']]),
      Location:             str(row[ci['Location']]),
      InfractionType:       str(row[ci['InfractionType']]),
      Severity:             str(row[ci['Severity']]),
      PointValue:           num(row[ci['PointValue']]),
      PointsBeforeReferral: num(row[ci['PointsBeforeReferral']]),
      PointsAfterReferral:  num(row[ci['PointsAfterReferral']]),
      Description:                str(row[ci['Description']]),
      IncludeDescriptionInEmail:  str(row[ci['IncludeDescriptionInEmail']]),
      TeacherName:          str(row[ci['TeacherName']]),
      ParentNotified:       str(row[ci['ParentNotified']]),
      TeacherNotified:      str(row[ci['TeacherNotified']]),
      Status:               str(row[ci['Status']]) || 'Open',
      AdminNotes:           str(row[ci['AdminNotes']]),
      TimestampFormatted:   (ts instanceof Date)
        ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MM-dd-yyyy h:mm a')
        : str(ts)
    });
  }

  referrals.sort(function(a, b) {
    return a.IncidentDate > b.IncidentDate ? -1 :
           a.IncidentDate < b.IncidentDate ?  1 : 0;
  });

  var ptSorted = referrals.slice().sort(function(a, b) {
    return a.IncidentDate < b.IncidentDate ? -1 :
           a.IncidentDate > b.IncidentDate ?  1 : 0;
  });
  var pointTimeline = ptSorted.map(function(r) {
    return {
      date:   r.IncidentDate,
      points: parseInt(r.PointsAfterReferral, 10) || 0,
      type:   r.InfractionType,
      delta:  parseInt(r.PointValue, 10) || 0
    };
  });

  var infMap = {};
  referrals.forEach(function(r) {
    if (r.InfractionType) infMap[r.InfractionType] = (infMap[r.InfractionType] || 0) + 1;
  });
  var infractionBreakdown = Object.keys(infMap)
    .map(function(k) { return { name: k, count: infMap[k] }; })
    .sort(function(a, b) { return b.count - a.count; });

  var nwSummary = cfg.nineWeeks.map(function(nw) {
    var count = 0;
    referrals.forEach(function(r) {
      if (r.IncidentDate >= nw.start && r.IncidentDate <= nw.end) count++;
    });
    return { label: nw.label, start: nw.start, end: nw.end, count: count };
  });

  var sumTotal = 0, sumOpen = 0, sumNegative = 0, sumPos = 0;
  referrals.forEach(function(r) {
    sumTotal++;
    if (r.Status !== 'Resolved') sumOpen++;
    if (r.Severity === 'Major' || r.Severity === 'Minor') sumNegative++;
    if (r.Severity === 'Positive') sumPos++;
  });

  return {
    student:             student,
    referrals:           referrals,
    pointTimeline:       pointTimeline,
    infractionBreakdown: infractionBreakdown,
    nwSummary:           nwSummary,
    nineWeeks:           cfg.nineWeeks,
    currentNineWeeks:    cfg.currentNineWeeks,
    semesterPoints:      cfg.semesterPoints,
    schoolName:          cfg.schoolName,
    pointTiers:          cfg.pointTiers,
    user:                user,
    summary: {
      total:    sumTotal,
      open:     sumOpen,
      negative: sumNegative,
      positive: sumPos
    }
  };
}

function getAllStudents() {
  var user = getCurrentUser();
  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied.');
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var cfg     = getConfig();
  var sp      = cfg.semesterPoints;
  var results = [];

  for (var i = 1; i < stuData.length; i++) {
    var pts = parseInt(stuData[i][STU_COL_POINTS], 10);
    if (isNaN(pts)) pts = sp;
    var pct = sp > 0 ? Math.round(pts / sp * 100) : 0;
    results.push({
      studentId:     stuData[i][STU_COL_ID].toString(),
      studentName:   displayName(stuData[i][STU_COL_FIRST], stuData[i][STU_COL_LAST]),
      grade:         stuData[i][STU_COL_GRADE].toString(),
      currentPoints: pts,
      pct:           pct,
      pointColor:    getPointColor(pct, cfg.pointTiers)
    });
  }
  results.sort(function(a, b) { return a.studentName.localeCompare(b.studentName); });

  return {
    students:       results,
    semesterPoints: sp,
    schoolName:     cfg.schoolName,
    pointTiers:     cfg.pointTiers,
    user:           user
  };
}

// =============================================================
// REPORT DATA
// =============================================================

function getReportData(filters) {
  var user = getCurrentUser();
  if (user.role !== 'admin' && user.role !== 'teacher') {
    throw new Error('Access denied.');
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
  var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
  var cfg     = getConfig();

  var hdrs = (refData && refData.length > 0) ? refData[0] : [];
  var ci   = {};
  for (var h = 0; h < hdrs.length; h++) {
    if (hdrs[h]) ci[hdrs[h].toString().trim()] = h;
  }

  var ptsMap = {};
  for (var s = 1; s < stuData.length; s++) {
    var sid = stuData[s][STU_COL_ID].toString();
    var pts = parseInt(stuData[s][STU_COL_POINTS], 10);
    ptsMap[sid] = isNaN(pts) ? cfg.semesterPoints : pts;
  }

  var rows       = [];
  var teacherSet = {};
  var infraSet   = {};

  for (var r = 1; r < refData.length; r++) {
    var row  = refData[r];
    var tch  = row[ci['TeacherName']]    ? row[ci['TeacherName']].toString()    : '';
    var inf  = row[ci['InfractionType']] ? row[ci['InfractionType']].toString() : '';
    var ts   = row[ci['Timestamp']];
    var sid2 = row[ci['StudentID']]      ? row[ci['StudentID']].toString()      : '';

    if (tch) teacherSet[tch] = true;
    if (inf) infraSet[inf]   = true;

    var curPts = ptsMap[sid2] !== undefined ? ptsMap[sid2] : cfg.semesterPoints;
    var curPct = cfg.semesterPoints > 0 ? Math.round(curPts / cfg.semesterPoints * 100) : 0;

    rows.push({
      ID:                   row[ci['ID']] !== undefined ? row[ci['ID']].toString() : '',
      IncidentDate:         formatDateStr(row[ci['IncidentDate']]),
      IncidentTime:         formatTimeStr(row[ci['IncidentTime']]),
      StudentID:            sid2,
      StudentName:          row[ci['StudentName']] ? row[ci['StudentName']].toString() : '',
      Grade:                row[ci['Grade']]       ? row[ci['Grade']].toString()       : '',
      InfractionType:       inf,
      Severity:             row[ci['Severity']]    ? row[ci['Severity']].toString()    : '',
      PointValue:           parseFloat(row[ci['PointValue']]) || 0,
      PointsBeforeReferral: parseFloat(row[ci['PointsBeforeReferral']]) || 0,
      PointsAfterReferral:  parseFloat(row[ci['PointsAfterReferral']]) || 0,
      Location:             row[ci['Location']]    ? row[ci['Location']].toString()    : '',
      Description:                row[ci['Description']] ? row[ci['Description']].toString() : '',
      IncludeDescriptionInEmail:  row[ci['IncludeDescriptionInEmail']] ? row[ci['IncludeDescriptionInEmail']].toString() : '',
      TeacherName:          tch,
      ParentNotified:       row[ci['ParentNotified']]  ? row[ci['ParentNotified']].toString()  : '',
      TeacherNotified:      row[ci['TeacherNotified']] ? row[ci['TeacherNotified']].toString() : '',
      Status:               row[ci['Status']]      ? row[ci['Status']].toString()      : 'Open',
      AdminNotes:           row[ci['AdminNotes']]  ? row[ci['AdminNotes']].toString()  : '',
      TimestampFormatted:   (ts instanceof Date)
        ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MM-dd-yyyy h:mm a')
        : '',
      CurrentPoints:        curPts,
      CurrentPointsColor:   getPointColor(curPct, cfg.pointTiers)
    });
  }

  rows.sort(function(a, b) {
    return a.IncidentDate > b.IncidentDate ? -1 :
           a.IncidentDate < b.IncidentDate ?  1 : 0;
  });

  return {
    rows:              rows,
    headers:           REFERRAL_HEADERS,
    teacherOptions:    Object.keys(teacherSet).sort(),
    infractionOptions: Object.keys(infraSet).sort(),
    semesterPoints:    cfg.semesterPoints,
    schoolName:        cfg.schoolName,
    pointTiers:        cfg.pointTiers,
    user:              user
  };
}

function updateReferralRow(referralId, newStatus, newAdminNotes) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REFERRALS);
  var data   = sheet.getDataRange().getValues();
  var idCol  = REFERRAL_HEADERS.indexOf('ID');
  var stCol  = REFERRAL_HEADERS.indexOf('Status');
  var anCol  = REFERRAL_HEADERS.indexOf('AdminNotes');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == referralId) {
      if (stCol >= 0) sheet.getRange(i + 1, stCol + 1).setValue(newStatus);
      if (anCol >= 0) sheet.getRange(i + 1, anCol + 1).setValue(newAdminNotes);
      return { success: true };
    }
  }
  return { success: false, error: 'Referral #' + referralId + ' not found.' };
}

function resetSemesterPoints() {
  try {
    var user = getCurrentUser();
    if (!user || user.role !== 'admin') {
      return { success: false, error: 'Admin access required.' };
    }
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_STUDENTS);
    var data  = sheet.getDataRange().getValues();
    var cfg   = getConfig();
    var now   = new Date();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, STU_COL_POINTS + 1).setValue(cfg.semesterPoints);
      sheet.getRange(i + 1, STU_COL_POINTS_DATE + 1).setValue(now);
      count++;
    }
    SpreadsheetApp.flush();
    return { success: true, count: count };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================
// UTILITIES
// =============================================================

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  }
}

function getLastId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  var idCol = REFERRAL_HEADERS.indexOf('ID') + 1;
  var ids   = sheet.getRange(2, idCol, lastRow - 1, 1).getValues()
                   .map(function(row) { return row[0]; })
                   .filter(function(v) { return typeof v === 'number' && v > 0; });
  return ids.length > 0 ? Math.max.apply(null, ids) : 0;
}

// =============================================================
// GUID GENERATOR
// =============================================================

/**
 * Generates a simple collision-resistant GUID prefixed with 'pg-'
 * (parent guid) for readability in the sheet.
 * Uses Math.random() which is sufficient for non-cryptographic
 * unique identifiers in a school-scale dataset.
 */
function generateGuid() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var guid  = 'pg-';
  for (var i = 0; i < 12; i++) {
    guid += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return guid;
}

// =============================================================
// STAFF MANAGEMENT
// =============================================================

/**
 * Returns all staff rows for the Staff manager UI.
 * Admin only.
 */
function getStaffList() {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STAFF);
  var data  = sheet.getDataRange().getValues();
  var rows  = [];

  for (var i = 1; i < data.length; i++) {
    var first = data[i][STAFF_COL_FIRST] ? data[i][STAFF_COL_FIRST].toString().trim() : '';
    var last  = data[i][STAFF_COL_LAST]  ? data[i][STAFF_COL_LAST].toString().trim()  : '';
    var email = data[i][STAFF_COL_EMAIL] ? data[i][STAFF_COL_EMAIL].toString().trim() : '';
    var role  = data[i][STAFF_COL_ROLE]  ? data[i][STAFF_COL_ROLE].toString().trim().toLowerCase() : '';
    if (!email) continue;
    rows.push({ firstName: first, lastName: last, email: email, role: role });
  }

  return { success: true, rows: rows };
}

/**
 * Replaces the entire Staff sheet content with the provided rows.
 * Invalidates CacheService for any email that changed role or was removed
 * so the role change takes effect on the user's next page load.
 * Admin only.
 */
function saveStaff(rows) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    return { success: false, error: 'At least one staff member is required.' };
  }

  // Validate — must have at least one admin
  var hasAdmin = rows.some(function(r) { return r.role === 'admin'; });
  if (!hasAdmin) {
    return { success: false, error: 'At least one staff member must have the admin role.' };
  }

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_STAFF);

    // Read existing rows to detect role changes for cache invalidation
    var existing  = sheet.getDataRange().getValues();
    var oldRoles  = {};
    for (var e = 1; e < existing.length; e++) {
      var eEmail = existing[e][STAFF_COL_EMAIL] ? existing[e][STAFF_COL_EMAIL].toString().trim().toLowerCase() : '';
      var eRole  = existing[e][STAFF_COL_ROLE]  ? existing[e][STAFF_COL_ROLE].toString().trim().toLowerCase()  : '';
      if (eEmail) oldRoles[eEmail] = eRole;
    }

    // Clear data rows (keep header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }

    // Write new rows
    var writeData = rows.map(function(r) {
      return [
        sanitizeText(r.firstName  || ''),
        sanitizeText(r.lastName   || ''),
        sanitizeText(r.email).toLowerCase(),
        r.role === 'admin' ? 'admin' : 'teacher'
      ];
    });
    sheet.getRange(2, 1, writeData.length, 4).setValues(writeData);
    SpreadsheetApp.flush();

    // Invalidate cache for anyone whose role changed or was removed
    var newEmails = {};
    rows.forEach(function(r) { newEmails[r.email.toLowerCase()] = r.role; });

    // Changed role or removed
    Object.keys(oldRoles).forEach(function(email) {
      if (!newEmails[email] || newEmails[email] !== oldRoles[email]) {
        invalidateUserCache(email);
      }
    });
    // Newly added (no cache to invalidate but harmless to call)
    Object.keys(newEmails).forEach(function(email) {
      if (!oldRoles[email]) {
        invalidateUserCache(email);
      }
    });

    // Also clear nav cache key so Settings link updates on next load
    // (can't clear client sessionStorage from server — handled client-side
    // by Settings.html after a successful saveStaff response)

    _user = null; // reset execution cache

    return { success: true, count: writeData.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================
// INFRACTIONS MANAGEMENT
// =============================================================

/**
 * Returns all infractions rows for the Infractions manager UI.
 * Admin only.
 */
function getInfractionsList() {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_INFRACTIONS);
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0].map(function(h) { return h.toString().trim(); });
  var ni    = hdrs.indexOf(INF_COL_NAME);
  var pi    = hdrs.indexOf(INF_COL_POINTS);
  var si    = hdrs.indexOf(INF_COL_SEVERITY);
  var nti   = hdrs.indexOf(INF_COL_NOTES);
  var rows  = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][ni] ? data[i][ni].toString().trim() : '';
    if (!name) continue;
    var pts = parseInt(data[i][pi], 10);
    rows.push({
      name:       name,
      pointValue: isNaN(pts) ? 0 : pts,
      severity:   si  >= 0 ? data[i][si].toString().trim()  : '',
      notes:      nti >= 0 ? data[i][nti].toString().trim() : ''
    });
  }

  return { success: true, rows: rows };
}

/**
 * Replaces the entire Infractions sheet content with the provided rows.
 * Clears the execution-level _infs cache so getInfractions() re-reads.
 * Admin only.
 */
function saveInfractions(rows) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    return { success: false, error: 'At least one infraction is required.' };
  }

  // Validate each row. Severity is optional — an empty string means
  // "None" (no severity assigned). Major/Minor/Positive remain the
  // only non-empty values allowed.
  var VALID_SEVERITIES = ['Major', 'Minor', 'Positive', ''];
  for (var v = 0; v < rows.length; v++) {
    var r = rows[v];
    if (!r.name || r.name.toString().trim() === '') {
      return { success: false, error: 'Row ' + (v + 1) + ': Infraction name is required.' };
    }
    if (isNaN(parseInt(r.pointValue, 10))) {
      return { success: false, error: 'Row ' + (v + 1) + ': Point value must be a number.' };
    }
    if (VALID_SEVERITIES.indexOf((r.severity || '').toString().trim()) < 0) {
      return { success: false, error: 'Row ' + (v + 1) + ': Severity must be Major, Minor, Positive, or left blank for None.' };
    }
  }

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_INFRACTIONS);

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }

    var writeData = rows.map(function(r) {
      return [
        sanitizeText(r.name),
        parseInt(r.pointValue, 10),
        r.severity,
        sanitizeText(r.notes || '')
      ];
    });
    sheet.getRange(2, 1, writeData.length, 4).setValues(writeData);
    SpreadsheetApp.flush();

    _infs = null; // clear execution cache

    return { success: true, count: writeData.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================
// STUDENT ROSTER MANAGEMENT
// =============================================================

/**
 * Returns all students for the roster manager UI.
 * Admin only.
 */
function getStudentsList() {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEET_STUDENTS);
  var data    = sheet.getDataRange().getValues();
  var cfg     = getConfig();
  var rows    = [];

  for (var i = 1; i < data.length; i++) {
    var id = data[i][STU_COL_ID] ? data[i][STU_COL_ID].toString().trim() : '';
    if (!id) continue;
    var pts = parseInt(data[i][STU_COL_POINTS], 10);
    rows.push({
      studentId:     id,
      studentName:   displayName(
        data[i][STU_COL_FIRST] ? data[i][STU_COL_FIRST].toString().trim() : '',
        data[i][STU_COL_LAST]  ? data[i][STU_COL_LAST].toString().trim()  : ''
      ),
      grade:         data[i][STU_COL_GRADE] ? data[i][STU_COL_GRADE].toString().trim() : '',
      currentPoints: isNaN(pts) ? cfg.semesterPoints : pts
    });
  }

  rows.sort(function(a, b) { return a.studentName.localeCompare(b.studentName); });
  return { success: true, rows: rows, semesterPoints: cfg.semesterPoints };
}

/**
 * Adds a new student or updates an existing one.
 * - New student: StudentID must not already exist; CurrentPoints
 *   defaults to SemesterStartPoints.
 * - Existing student: finds the row by StudentID and updates name/grade.
 * Admin only.
 */
function saveStudent(student) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var id     = sanitizeText(student.studentId  || '').toUpperCase();
  var first  = sanitizeText(student.firstName   || '');
  var last   = sanitizeText(student.lastName    || '');
  var middle = sanitizeText(student.middleName  || '');
  var grade  = sanitizeText(student.grade        || '');

  if (!id)    return { success: false, error: 'Student ID is required.' };
  if (!first) return { success: false, error: 'First name is required.' };
  if (!last)  return { success: false, error: 'Last name is required.' };
  if (!grade) return { success: false, error: 'Grade is required.' };

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName(SHEET_STUDENTS);
    var data    = sheet.getDataRange().getValues();
    var cfg     = getConfig();
    var now     = new Date();

    // Check for existing row
    for (var i = 1; i < data.length; i++) {
      if (data[i][STU_COL_ID].toString().trim().toUpperCase() === id) {
        // Update existing — preserve points
        sheet.getRange(i + 1, STU_COL_FIRST  + 1).setValue(first);
        sheet.getRange(i + 1, STU_COL_LAST   + 1).setValue(last);
        sheet.getRange(i + 1, STU_COL_MIDDLE + 1).setValue(middle);
        sheet.getRange(i + 1, STU_COL_GRADE  + 1).setValue(grade);
        SpreadsheetApp.flush();
        return { success: true, action: 'updated' };
      }
    }

    // New student — use override points if provided, else semester default
    var pts = (student.currentPoints !== undefined && student.currentPoints !== '')
      ? parseInt(student.currentPoints, 10)
      : cfg.semesterPoints;
    if (isNaN(pts) || pts < 0) pts = cfg.semesterPoints;

    sheet.appendRow([id, first, last, middle, grade, pts, now]);
    SpreadsheetApp.flush();
    return { success: true, action: 'added' };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Deletes a student row by StudentID.
 * Blocked if the student has any referral records — referral history
 * must be preserved even if a student leaves.
 * Admin only.
 */
function deleteStudent(studentId) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var id = studentId.toString().trim();

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
    var sidIdx  = REFERRAL_HEADERS.indexOf('StudentID');

    // Block deletion if referrals exist
    for (var r = 1; r < refData.length; r++) {
      if ((refData[r][sidIdx] || '').toString().trim() === id) {
        return {
          success: false,
          error: 'Cannot delete — this student has referral records. ' +
                 'Referral history must be preserved.'
        };
      }
    }

    var stuSheet = ss.getSheetByName(SHEET_STUDENTS);
    var stuData  = stuSheet.getDataRange().getValues();
    for (var i = 1; i < stuData.length; i++) {
      if (stuData[i][STU_COL_ID].toString().trim() === id) {
        stuSheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Student ID not found: ' + id };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Bulk imports students from a CSV-parsed array.
 * Each row: { studentId, studentName, grade }
 * - Skips rows with duplicate StudentIDs (existing or within import).
 * - CurrentPoints defaults to SemesterStartPoints for all new students.
 * - Returns counts of added, skipped, and any errors.
 * Admin only.
 */
function importStudents(rows) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    return { success: false, error: 'No rows provided.' };
  }

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName(SHEET_STUDENTS);
    var data    = sheet.getDataRange().getValues();
    var cfg     = getConfig();
    var now     = new Date();

    // Build set of existing IDs for fast lookup
    var existingIds = {};
    for (var e = 1; e < data.length; e++) {
      var eid = data[e][STU_COL_ID].toString().trim().toUpperCase();
      if (eid) existingIds[eid] = true;
    }

    var added    = 0;
    var skipped  = 0;
    var errors   = [];
    var seenInImport = {};

    for (var i = 0; i < rows.length; i++) {
      try {
        var r      = rows[i];
        var id     = sanitizeText(r.studentId  || '').toUpperCase();
        var first  = sanitizeText(r.firstName   || '');
        var last   = sanitizeText(r.lastName    || '');
        var middle = sanitizeText(r.middleName  || '');
        var grade  = sanitizeText(r.grade        || '');

        if (!id || !first || !last || !grade) {
          errors.push('Row ' + (i + 1) + ': missing required field (ID, first name, last name, or grade).');
          continue;
        }

        if (existingIds[id] || seenInImport[id]) {
          skipped++;
          continue;
        }

        sheet.appendRow([id, first, last, middle, grade, cfg.semesterPoints, now]);
        existingIds[id]  = true;
        seenInImport[id] = true;
        added++;

      } catch (rowErr) {
        errors.push('Row ' + (i + 1) + ': ' + rowErr.message);
      }
    }

    SpreadsheetApp.flush();
    return { success: true, added: added, skipped: skipped, errors: errors };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================
// PARENT CONTACTS MANAGEMENT
// =============================================================

/**
 * Searches existing parent contacts by name or email.
 * Returns unique parent records (deduplicated by ParentGUID)
 * for the search-before-create flow.
 * Admin only.
 */
function searchParentContacts(query) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var q = (query || '').toString().trim().toLowerCase();
  if (!q || q.length < 2) {
    return { success: true, results: [] };
  }

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var data  = ss.getSheetByName(SHEET_PARENT).getDataRange().getValues();
    var seen  = {}; // deduplicate by GUID
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var guid  = data[i][PARENT_COL_GUID]  ? data[i][PARENT_COL_GUID].toString().trim()  : '';
      var first = data[i][PARENT_COL_FIRST] ? data[i][PARENT_COL_FIRST].toString().trim() : '';
      var last  = data[i][PARENT_COL_LAST]  ? data[i][PARENT_COL_LAST].toString().trim()  : '';
      var email = data[i][PARENT_COL_EMAIL] ? data[i][PARENT_COL_EMAIL].toString().trim() : '';
      var phone = data[i][PARENT_COL_PHONE] ? data[i][PARENT_COL_PHONE].toString().trim() : '';
      var name  = displayName(first, last);

      if (!guid) continue;
      if (seen[guid]) continue;

      var matches = name.toLowerCase().indexOf(q) >= 0 ||
                    first.toLowerCase().indexOf(q) >= 0 ||
                    last.toLowerCase().indexOf(q) >= 0 ||
                    email.toLowerCase().indexOf(q) >= 0;

      if (matches) {
        seen[guid] = true;
        results.push({
          guid: guid, firstName: first, lastName: last,
          name: name, email: email, phone: phone
        });
      }
    }

    results.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return { success: true, results: results };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Returns the contact for a student and any siblings sharing the
 * same ParentGUID, for the sibling-checkbox UI on the profile page.
 * Admin only.
 */
function getStudentContacts(studentId) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var pData   = ss.getSheetByName(SHEET_PARENT).getDataRange().getValues();
    var stuData = ss.getSheetByName(SHEET_STUDENTS).getDataRange().getValues();
    var id      = studentId.toString().trim();

    // Find the contact row for this student
    var contact   = null;
    var parentGuid = '';

    for (var i = 1; i < pData.length; i++) {
      if (pData[i][PARENT_COL_STUDENT_ID].toString().trim() === id) {
        parentGuid = pData[i][PARENT_COL_GUID]  ? pData[i][PARENT_COL_GUID].toString().trim()  : '';
        var cFirst = pData[i][PARENT_COL_FIRST] ? pData[i][PARENT_COL_FIRST].toString().trim() : '';
        var cLast  = pData[i][PARENT_COL_LAST]  ? pData[i][PARENT_COL_LAST].toString().trim()  : '';
        contact = {
          studentId:   id,
          guid:        parentGuid,
          firstName:   cFirst,
          lastName:    cLast,
          parentName:  displayName(cFirst, cLast),
          parentEmail: pData[i][PARENT_COL_EMAIL] ? pData[i][PARENT_COL_EMAIL].toString().trim() : '',
          phone:       pData[i][PARENT_COL_PHONE] ? pData[i][PARENT_COL_PHONE].toString().trim() : ''
        };
        break;
      }
    }

    // Build a student name lookup for sibling display
    var stuNames = {};
    for (var s = 1; s < stuData.length; s++) {
      var sid = stuData[s][STU_COL_ID].toString().trim();
      stuNames[sid] = displayName(stuData[s][STU_COL_FIRST], stuData[s][STU_COL_LAST]);
    }

    // Find siblings sharing the same ParentGUID (excluding this student)
    var siblings = [];
    if (parentGuid) {
      for (var j = 1; j < pData.length; j++) {
        var rowGuid = pData[j][PARENT_COL_GUID] ? pData[j][PARENT_COL_GUID].toString().trim() : '';
        var rowSid  = pData[j][PARENT_COL_STUDENT_ID].toString().trim();
        if (rowGuid === parentGuid && rowSid !== id) {
          siblings.push({
            studentId:   rowSid,
            studentName: stuNames[rowSid] || rowSid
          });
        }
      }
    }

    return {
      success:  true,
      contact:  contact,   // null if no contact on file
      siblings: siblings   // [] if no siblings
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Saves or updates the parent contact for a student.
 * If updateSiblingIds is a non-empty array, updates those sibling rows too.
 * If contact.guid is empty, generates a new ParentGUID.
 * If contact.guid is provided (from search-before-create), reuses it.
 * Admin only.
 */
function saveParentContact(studentId, contact, updateSiblingIds) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var id    = studentId.toString().trim();
  if (!id) return { success: false, error: 'Student ID is required.' };

  var first = sanitizeText(contact.firstName   || '');
  var last  = sanitizeText(contact.lastName    || '');
  var email = sanitizeText(contact.parentEmail || '').toLowerCase();
  var phone = sanitizeText(contact.phone        || '');
  var guid  = contact.guid ? contact.guid.toString().trim() : generateGuid();

  if (!first) return { success: false, error: 'Parent first name is required.' };
  if (!last)  return { success: false, error: 'Parent last name is required.' };
  if (!email) return { success: false, error: 'Parent email is required.' };

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PARENT);
    var data  = sheet.getDataRange().getValues();

    // Collect all student IDs to update (this student + selected siblings)
    var toUpdate = [id];
    if (Array.isArray(updateSiblingIds)) {
      updateSiblingIds.forEach(function(sid) {
        var s = sid.toString().trim();
        if (s && s !== id) toUpdate.push(s);
      });
    }

    var updatedRows = {};

    // Update existing rows for each target student
    for (var i = 1; i < data.length; i++) {
      var rowSid = data[i][PARENT_COL_STUDENT_ID].toString().trim();
      if (toUpdate.indexOf(rowSid) >= 0) {
        sheet.getRange(i + 1, PARENT_COL_GUID  + 1).setValue(guid);
        sheet.getRange(i + 1, PARENT_COL_FIRST + 1).setValue(first);
        sheet.getRange(i + 1, PARENT_COL_LAST  + 1).setValue(last);
        sheet.getRange(i + 1, PARENT_COL_EMAIL + 1).setValue(email);
        sheet.getRange(i + 1, PARENT_COL_PHONE + 1).setValue(phone);
        updatedRows[rowSid] = true;
      }
    }

    // Append new rows for any target student not yet in the sheet
    toUpdate.forEach(function(sid) {
      if (!updatedRows[sid]) {
        sheet.appendRow([sid, guid, first, last, email, phone]);
      }
    });

    SpreadsheetApp.flush();
    return { success: true, guid: guid, updated: toUpdate.length };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Removes the parent contact row for a single student.
 * Does not affect sibling rows — those retain their contact info.
 * Admin only.
 */
function deleteParentContact(studentId) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }

  var id = studentId.toString().trim();

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PARENT);
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][PARENT_COL_STUDENT_ID].toString().trim() === id) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'No contact found for student: ' + id };

  } catch (e) {
    return { success: false, error: e.message };
  }
}