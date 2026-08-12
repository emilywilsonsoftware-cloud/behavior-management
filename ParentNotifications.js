/**
 * ── Parent Notifications — manual mail-merge fallback ──────────────────
 *
 * Supports sending notifications manually (Word Mail Merge, run by an
 * admin from their own email) instead of GmailApp.sendEmail(), which is
 * being silently filtered by the district's Exchange/Defender setup.
 *
 * Verified against the actual Code.js: uses the real REFERRAL_HEADERS /
 * PARENT_COL_* constants and SHEET_REFERRALS / SHEET_PARENT names, and
 * mirrors sendDailyParentEmails()'s exact filtering logic (future-dated
 * exclusion, 'Yes'/'N/A' exclusion, positive-note exclusion) so this
 * list and the daily-digest logic never disagree about what's pending.
 */

/**
 * Returns one entry per (student, contact) pair with at least one
 * pending demerit referral — mirrors sendDailyParentEmails()'s grouping
 * and filtering exactly, but returns data instead of sending email.
 * Includes every contact type on file for that student (Parent/Guardian,
 * Administrator, Counselor, Case Manager), same as the daily digest.
 *
 * Admin-only.
 */
/**
 * Read-only status check for the Parent Notifications panel — lets it
 * show whether the automated (GmailApp) path is still enabled in
 * Settings, without needing a second toggle that could drift out of
 * sync with EmailNotificationsEnabled. Admin-only, same as everything
 * else in this file.
 */
function getEmailModeStatus() {
  var user = getCurrentUser();
  if (!user || user.role !== 'admin') {
    throw new Error('Admin access required.');
  }
  var cfg = getConfig();
  return { automatedEnabled: !!cfg.emailEnabled };
}

function getPendingParentNotifications() {
  var user = getCurrentUser();
  if (!user || user.role !== 'admin') {
    throw new Error('Admin access required.');
  }

  var cfg   = getConfig();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var data  = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
  var today = formatDate(new Date());

  var idIdx    = REFERRAL_HEADERS.indexOf('ID');
  var sidIdx   = REFERRAL_HEADERS.indexOf('StudentID');
  var snIdx    = REFERRAL_HEADERS.indexOf('StudentName');
  var grIdx    = REFERRAL_HEADERS.indexOf('Grade');
  var dtIdx    = REFERRAL_HEADERS.indexOf('IncidentDate');
  var tmIdx    = REFERRAL_HEADERS.indexOf('IncidentTime');
  var locIdx   = REFERRAL_HEADERS.indexOf('Location');
  var infIdx   = REFERRAL_HEADERS.indexOf('InfractionType');
  var ptIdx    = REFERRAL_HEADERS.indexOf('PointValue');
  var afIdx    = REFERRAL_HEADERS.indexOf('PointsAfterReferral');
  var tchIdx   = REFERRAL_HEADERS.indexOf('TeacherName');
  var descIdx  = REFERRAL_HEADERS.indexOf('Description');
  var incEmIdx = REFERRAL_HEADERS.indexOf('IncludeDescriptionInEmail');
  var pnIdx    = REFERRAL_HEADERS.indexOf('ParentNotified');

  // studentId -> [{ firstName, lastName, email }] — same shape as
  // sendDailyParentEmails()'s contactMap, all contact Types included.
  var contacts   = ss.getSheetByName(SHEET_PARENT).getDataRange().getValues();
  var contactMap = {};
  for (var p = 1; p < contacts.length; p++) {
    var sid   = contacts[p][PARENT_COL_STUDENT_ID] ? contacts[p][PARENT_COL_STUDENT_ID].toString().trim() : '';
    var email = contacts[p][PARENT_COL_EMAIL]      ? contacts[p][PARENT_COL_EMAIL].toString().trim()      : '';
    if (!sid || !email) continue;
    if (!contactMap[sid]) contactMap[sid] = [];
    contactMap[sid].push({
      firstName: contacts[p][PARENT_COL_FIRST] ? contacts[p][PARENT_COL_FIRST].toString().trim() : '',
      lastName:  contacts[p][PARENT_COL_LAST]  ? contacts[p][PARENT_COL_LAST].toString().trim()  : '',
      email:     email
    });
  }

  // studentId -> { studentName, grade, refs: [...] }
  var byStudent = {};
  for (var r = 1; r < data.length; r++) {
    var row      = data[r];
    var incDate  = formatDateStr(row[dtIdx]);
    var notified = row[pnIdx] ? row[pnIdx].toString() : '';

    if (incDate > today) continue;               // matches sendDailyParentEmails()
    if (notified === 'Yes' || notified === 'N/A') continue;

    var pts = parseInt(row[ptIdx], 10) || 0;
    if (pts > 0) continue; // positive notes never generate a parent notification — matches
                            // sendDailyParentEmails()'s exact check; note this means a
                            // 0-point infraction IS included, same as the real digest.

    var studentId = row[sidIdx] ? row[sidIdx].toString() : '';
    if (!studentId) continue;

    if (!byStudent[studentId]) {
      byStudent[studentId] = {
        studentName: row[snIdx],
        grade:       row[grIdx],
        refs:        []
      };
    }

    var includeDesc = (row[incEmIdx] || '').toString().trim().toLowerCase() === 'yes';

    byStudent[studentId].refs.push({
      id:             row[idIdx],
      incidentDate:   incDate,
      incidentTime:   formatTime12h(formatTimeStr(row[tmIdx])),
      location:       row[locIdx] ? row[locIdx].toString() : '',
      infractionType: row[infIdx] ? row[infIdx].toString() : '',
      pointValue:     pts,
      pointsAfter:    row[afIdx] !== undefined ? row[afIdx] : '',
      teacherName:    row[tchIdx] ? row[tchIdx].toString() : '',
      description:    includeDesc && row[descIdx] ? row[descIdx].toString() : ''
    });
  }

  var results = [];
  Object.keys(byStudent).forEach(function(sid) {
    var stu = byStudent[sid];
    var studentContacts = contactMap[sid];
    // No contact on file at all — silently skipped, same as
    // sendDailyParentEmails(). Worth periodically auditing separately.
    if (!studentContacts || studentContacts.length === 0) return;

    var displayNameForParent = parentSafeDisplayName_(stu.studentName);
    var subject = 'Behavior Notification \u2014 ' + cfg.schoolName + ' \u2014 ' + displayNameForParent;
    var body    = buildNotificationBody_(cfg, displayNameForParent, stu.grade, stu.refs);
    var referralIds = stu.refs.map(function(ref) { return ref.id; });

    studentContacts.forEach(function(contact) {
      results.push({
        key:            sid + ':' + contact.email,
        studentId:      sid,
        studentDisplay: displayNameForParent,
        referralIds:    referralIds,
        parentName:     displayName(contact.firstName, contact.lastName),
        parentEmail:    contact.email,
        subject:        subject,
        body:           body
      });
    });
  });

  return results;
}

/**
 * "First Last" (the actual StudentName format, per displayName() used
 * at submit time) -> "First L." for parent-facing display — full first
 * name plus last initial, chosen specifically to avoid the collision
 * that "F. Last" has when siblings share a last name.
 */
function parentSafeDisplayName_(storedName) {
  var words = (storedName || '').toString().trim().split(/\s+/);
  if (words.length < 2) return storedName || '';
  return words[0] + ' ' + words[words.length - 1].charAt(0) + '.';
}

function buildNotificationBody_(cfg, displayNameForParent, grade, refs) {
  var lines = [];

  lines.push('This is an automated behavior notification from ' + cfg.schoolName + ' for ' +
    displayNameForParent + (displayNameForParent.slice(-1) === '.' ? '' : '.'));
  lines.push('');
  lines.push(refs.length + ' Referral' + (refs.length !== 1 ? 's' : '') + ' Received');
  lines.push('');

  refs.forEach(function(ref, idx) {
    var ptsStr = ref.pointValue > 0 ? '+' + ref.pointValue : ref.pointValue.toString();
    lines.push('Referral ' + (idx + 1) + ':');
    lines.push(padLabel_('Date:') + ref.incidentDate);
    lines.push(padLabel_('Time:') + ref.incidentTime);
    lines.push(padLabel_('Location:') + ref.location);
    lines.push(padLabel_('Infraction:') + ref.infractionType);
    lines.push(padLabel_('Teacher:') + ref.teacherName);
    lines.push(padLabel_('Points:') + ptsStr + '  (Balance: ' + ref.pointsAfter + ' pts)');
    if (ref.description) lines.push(padLabel_('Notes:') + ref.description);
    lines.push('');
    lines.push('');
  });

  lines.push('Current Point Balance: ' + refs[refs.length - 1].pointsAfter + ' pts');
  lines.push('');

  // Pulled from Config's EmailFooterText (same field Settings and the
  // automated sendDailyParentEmails() digest already use) rather than
  // hardcoded here — keeps Settings the one real source of truth for
  // this text across both the automated and manual-export paths.
  lines.push(cfg.emailFooter);

  // Real newlines — Word/Outlook generally preserve embedded line
  // breaks from a quoted CSV field when merging to email, which is
  // what gives this layout its actual line breaks and alignment.
  // Alignment assumes a monospace-ish rendering, which is the common
  // default for plain-text email bodies (what OutMail.Body sets) —
  // worth a quick visual check via Mailings > Preview Results, or a
  // test send to yourself, since exact rendering varies by client.
  return lines.join('\n');
}

// Pads a field label (e.g. "Date:") to a fixed column width so the
// values line up underneath each other.
function padLabel_(label) {
  var LABEL_WIDTH = 13;
  var padded = label;
  while (padded.length < LABEL_WIDTH) padded += ' ';
  return padded;
}

/**
 * Marks the given referral IDs ParentNotified = 'Yes'. Admin-only.
 * Called automatically right after a CSV export (see
 * getPendingParentNotifications() callers in Report.html) — export IS
 * the point of "notification," by design: once the CSV is downloaded,
 * these referrals are treated as handled and won't be re-offered for
 * export tomorrow, whether or not the actual send later succeeds. The
 * counterpart to this is setParentNotifiedStatus() below, which lets an
 * admin manually revert one back to "not notified" for the rare case
 * where a send is known to have failed and needs to be re-sent.
 */
function markReferralsNotified(referralIds) {
  var user = getCurrentUser();
  if (!user || user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }
  if (!referralIds || referralIds.length === 0) {
    return { success: false, error: 'No referral IDs provided.' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REFERRALS);
  var data  = sheet.getDataRange().getValues();
  var idCol = REFERRAL_HEADERS.indexOf('ID');
  var pnCol = REFERRAL_HEADERS.indexOf('ParentNotified');

  var idSet = {};
  referralIds.forEach(function(id) { idSet[id.toString()] = true; });

  var updated = 0;
  for (var r = 1; r < data.length; r++) {
    if (idSet[data[r][idCol].toString()]) {
      sheet.getRange(r + 1, pnCol + 1).setValue('Yes');
      updated++;
    }
  }

  return { success: true, updated: updated };
}

/**
 * Manually sets a single referral's ParentNotified value — the escape
 * hatch for the rare case where a mail-merge send is known to have
 * failed (or was sent in error) after the CSV export already marked it
 * 'Yes'. Setting it back to 'No' makes it reappear in the Parent
 * Notifications panel so it gets exported/sent again next time.
 * Admin-only. Not wired into the change-log system used elsewhere in
 * Report.html (getChangeLog()) — that write path wasn't available to
 * verify against, so this change won't show up there yet.
 */
function setParentNotifiedStatus(referralId, newStatus) {
  var user = getCurrentUser();
  if (!user || user.role !== 'admin') {
    return { success: false, error: 'Admin access required.' };
  }
  if (newStatus !== 'Yes' && newStatus !== 'No') {
    return { success: false, error: 'newStatus must be "Yes" or "No".' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REFERRALS);
  var data  = sheet.getDataRange().getValues();
  var idCol = REFERRAL_HEADERS.indexOf('ID');
  var pnCol = REFERRAL_HEADERS.indexOf('ParentNotified');

  for (var r = 1; r < data.length; r++) {
    if (data[r][idCol].toString() === referralId.toString()) {
      sheet.getRange(r + 1, pnCol + 1).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false, error: 'Referral #' + referralId + ' not found.' };
}