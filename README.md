# BehaviorTracker — Google Apps Script Web App
### Student Behavior Referral Tracking on Google Workspace for Education
**Version 4.1 | Full System**

---

## Overview

BehaviorTracker is a complete, Google-native behavior management system built on Google Apps Script. It runs entirely inside your school's existing Google Workspace for Education — no software to install, no cloud subscriptions, no monthly fees.

### Pages

| URL Parameter | Page | Who Uses It |
|---|---|---|
| *(none)* or `?page=dashboard` | **Dashboard** — homepage with stats, charts, at-risk students | All staff |
| `?page=form` | **Referral Form** — submit behavior referrals | Teachers, Admins |
| `?page=positive` | **Positive Note** — award positive points (Write Off, Saturday School, etc.) | Admins |
| `?page=report` | **Admin Report** — filterable, paginated referral log | Admins, Teachers |
| `?page=student&id=S001` | **Student Profile** — individual history, point timeline, and the Point Range tool | Admins, Teachers |
| `?page=settings` | **Settings** — all school configuration | Admins |

---

## File Structure

```
BehaviorTrackerGAS/
├── Code.gs        — Server-side logic (routing, data, email, points)
├── Setup.gs       — One-time setup, admin spreadsheet menu, and triggers
├── Nav.html       — Shared navigation bar (included by all pages)
├── Dashboard.html — Homepage: stats, charts, at-risk, recent activity
├── Index.html     — Behavior referral entry form
├── Positive.html  — Positive note entry form (admin only)
├── Report.html    — Admin report — filters, pagination, Change Log, inline editing
├── Student.html   — Student profile, point timeline, and Point Range tool
├── Settings.html  — All school configuration (General, Behavior & Points, Staff, Students, Reset & Archive)
└── README.md      — This file
```

---

## Google Sheets Structure

| Sheet | Purpose | Who Edits |
|---|---|---|
| **Config** | School name, term start points, email settings, term start dates, dropdown lists (Locations, Contact Types), point-color tiers, positive points cap | Admin, via Settings |
| **Infractions** | All infraction types with point values and notes (positive point values = merit types, e.g. Write Off) | Admin, via Settings |
| **Referrals** | Master behavior log — auto-populated by the referral and positive note forms | Auto / Admin notes / Admin edits |
| **Students** | Student roster with live point balances | Admin, via Settings |
| **Staff** | Teacher/admin names, emails, and role | Admin, via Settings |
| **ParentContacts** | Referral email contacts per student (parents, guardians, counselors, etc.) — a student can have any number | Admin, via the Student page or when adding a student in Settings |
| **ChangeLog** | Audit trail of every referral that's been deleted, edited, or bulk-archived — who, when, and what changed | Auto — read-only in the app |

---

## Deployment Instructions

### Step 1 — Create the Spreadsheet
1. Sign in with your **school Google Workspace account**
2. Create a new Google Sheet — name it **BehaviorTracker**

### Step 2 — Add Script Files
1. Click **Extensions > Apps Script**
2. Create the following files and paste each file's contents:
   - Rename the default file to `Code.gs` — paste `Code.gs`
   - New Script file → name `Setup` → paste `Setup.gs`
   - New HTML file → name `Nav` → paste `Nav.html`
   - New HTML file → name `Dashboard` → paste `Dashboard.html`
   - New HTML file → name `Index` → paste `Index.html`
   - New HTML file → name `Positive` → paste `Positive.html`
   - New HTML file → name `Report` → paste `Report.html`
   - New HTML file → name `Student` → paste `Student.html`
   - New HTML file → name `Settings` → paste `Settings.html`
3. Save all files

### Step 3 — Run Initial Setup
1. In Apps Script, select function `setupSpreadsheet` → click **Run**
2. Approve Google permissions when prompted
3. Switch back to the spreadsheet — 7 sheets will be created: Config, Infractions, Referrals, Students, Staff, ParentContacts, and ChangeLog

### Step 4 — Configure Your Data

Almost everything below can now be managed directly from the **Settings page** in the web app once it's deployed (Step 5) — you generally won't need to edit the spreadsheet by hand except for the very first admin account.

**Staff sheet — set this first, directly in the sheet, since you need at least one admin before you can log into Settings:**
- Find the placeholder row (`Admin | User | admin@yourschool.edu | admin`) and replace the email with your actual Google Workspace login email. Leave `Role` as `admin` (not case-sensitive — "Admin" or "ADMIN" work identically).
- Role is looked up from this sheet, not from any setting in Config — there is no separate "admin list" elsewhere.

**Everything else, from Settings once deployed:**
- **General tab** — school name, term start points, email settings, Locations list, Contact Types list, and Term Start Dates
- **Behavior & Points tab** — point-balance color tiers, Infraction Types (with point values), Positive Note Types, and the Positive Points Cap
- **Staff tab** — add more teachers and admins; `StaffEmail` must exactly match each person's Google Workspace login email
- **Students tab** — add students one at a time (optionally with a referral email contact at the same time) or bulk-import a CSV; view the full roster
- **Reset & Archive tab** — reset all point balances for a new term, archive a single term's referrals to CSV without resetting anything, or archive-and-reset for a new school year

---

## Role System

Roles are determined automatically from the user's Google login email:

| Role | How Assigned | Access |
|---|---|---|
| **admin** | `Role` column in the Staff sheet set to `admin` | All pages, including Settings, Positive Note, and full Report features (Change Log, editing referrals, the Teacher filter) |
| **teacher** | `Role` column in the Staff sheet set to `teacher` | Dashboard, Referral Form, Report (view + Point Range tool), Student profiles |
| **viewer** | Any other domain user not listed in Staff at all | Dashboard, Report (view only) |

Role badge is shown in the top-right navigation bar. Admin-only nav links (Settings, Positive Note) and the Report page's Teacher filter are hidden by default and only revealed once the user's role is confirmed — non-admins never see them flash on screen.

---

## Points System

### How Points Work
- Every student starts each **term** with `TermStartPoints` (default: 100), configurable from Settings → General
- Each infraction type in the **Infractions** sheet has a `PointValue`
  - Negative values deduct points (behavior infractions)
  - Positive values add points (merit / positive behaviors — submitted from the Positive Note page, subject to a configurable per-term cap that admins can override with a reason)
- **Balances can go negative**, and do so deliberately — a student's true point total is always shown as-is, never artificially floored at 0. Hiding how far under 0 a student actually is would hide exactly the information an admin most needs to see; the Point Range tool on the Student page can filter down to negative balances specifically for this reason
- Points are updated automatically when a referral or positive note is submitted, when a referral is deleted, or when a referral's infraction type is edited — all of these go through Apps Script's `LockService`, so two people (e.g. two teachers submitting for the same student) can't race each other and silently overwrite one another's point changes
- `PointsLastUpdated` timestamp is recorded in the Students sheet

### Point Balance Color Coding
Colors and their thresholds are fully configurable from Settings → Behavior & Points (not fixed percentages) — the default is:

| Balance | Color | Meaning |
|---|---|---|
| ≥ 70% of starting points | Green | On track |
| 40–69% of starting points | Amber | Watch closely |
| < 40% of starting points | Red | At risk |

### Term Reset
Points reset **per term**, not per semester — most schools will do this 4 times a year rather than 2. To reset:
1. Go to **Settings → Reset & Archive → Reset Term Points**, or use the **BehaviorTracker Admin → Reset Term Points** spreadsheet menu item
2. Confirm — all students reset to `TermStartPoints`

> **Note:** This does not delete referral history. Only the live `CurrentPoints` balance is reset. For a full year-end reset that also archives and clears referral history, use **Start New Year** in the same tab.

### Archiving a Single Term's Referrals

For a school with an unusually heavy referral volume in one term, **Settings → Reset & Archive → Archive a Term's Referrals** is a lighter-weight alternative to Start New Year — it exports just one selected term's referrals to a downloaded CSV, then permanently removes only those rows from the Referrals sheet. **Student point balances are not changed** by this — it trims the size of the detailed log without undoing anything those referrals already affected on a student's balance. It's correct to run at any time, but lowest-risk right after that term's points have already been reset, since at that point the archived rows have no remaining connection to any live balance at all.

Every archive action — deletions, edits, and this kind of bulk archiving — is recorded once as a summary entry in the Change Log (see Admin Report Page, below), not once per referral.

---

## Contacts

Each student can have any number of referral email contacts — parents, guardians, counselors, case managers, or other staff — each tagged with a configurable **Contact Type** (Settings → General → Contact Types).

- **From the Student page:** open a student's profile → Edit Info → Contacts, to add, edit, or remove contacts. Adding one offers a search across both *existing contacts* (reuse someone already on file for another student, e.g. a sibling's parent) and *staff* (pull in an existing teacher/counselor by name instead of retyping their email).
- **From Settings, when adding a new student:** the Add Student form includes an optional, repeatable "Referral Email Contacts" section — add as many as needed at the same time the student is created, or skip it and add them later.

All of a student's contacts receive the same end-of-day referral email digest.

---

## Email Notifications

Every submitted referral sends up to two emails automatically:

**Parent/Contact Email** — batched once daily, sent to every contact on file for that student:
- Student name, grade, incident details
- Points deducted/added and new balance
- Teacher name and the school's name in both the subject and body
- School name and footer text from Config

**Teacher Confirmation** — sent immediately to the submitting teacher:
- Referral ID, student name, incident summary
- Point update with before/after balance

To disable all emails: turn off notifications from Settings → General, or set `EmailNotificationsEnabled` to `No` in the Config sheet directly.

---

## Dashboard

The Dashboard is the default homepage (`?page=dashboard`) and shows a different view depending on role.

**Admin view:**
- **8 stat cards:** total referrals this term, today, this week, major-point-deduction count, total students, at-risk count, term starting points, and more
- **Monthly trend chart:** referral volume over the trailing 12 months
- **Top infractions chart:** most common infraction types
- **At-risk students table:** the 10 lowest-balance students, linkable to profiles
- **Recent activity feed**
- **Location breakdown** and **teacher activity this month**

**Teacher view:** a lighter version scoped to the teacher's own submissions.

---

## Student Profile Page

Access via `?page=student&id=[StudentID]`, or by searching from the page directly. Has two tabs:

**Student Lookup** (default) — search by name or ID, filter by grade. Selecting a student shows:
- Name, grade, ID, live point balance with color-coded progress bar
- Contacts, with an Edit Info modal for managing student details and contacts
- Summary badges (negative / positive referral counts)
- Term-by-term referral history tabs, point balance timeline, and infraction breakdown

**Point Range** — a separate tool for finding students by point balance (and optionally grade), sortable, with a CSV download of the full matching list. Unlike the Report page's filters, this always includes every student, including anyone with a completely clean record and no referrals at all — it reads student balances directly rather than filtering referral rows.

---

## Admin Report Page

Access via `?page=report`. Loads fast even with a large referral history — the first page of results appears immediately while the rest loads in the background.

- **Filters:** student (type-ahead search), infraction type, term, date range, and — admin only — teacher
- **Summary cards:** affected students, negative referrals, positive notes, average points deducted
- **Sortable, paginated table**
- **Detail panel:** view full referral detail; admins can edit the incident's date, time, location, infraction type, and description, or delete the referral entirely (with a required reason)
- **Change Log** (admin only): every deletion, edit, and bulk archive action, showing who made the change, when, and what changed — a bulk archive gets one summary entry, not one per referral
- **Export CSV:** exports the complete currently-filtered set, not just the visible page

---

## Positive Note Page

Access via `?page=positive` (admin only). A dedicated form for awarding positive points — Write Off, Saturday School, or other configured positive infraction types — to one or more students at once. Submissions that would put a student over the configured per-term cap require confirmation before saving, but admins can always override.

---

## Settings Page

Access via `?page=settings` (admin only). Everything below is configured here — direct spreadsheet editing is rarely needed after initial setup.

- **General** — school name, term start points, email settings and footer, Locations, Contact Types, and Term Start Dates
- **Behavior & Points** — point balance color tiers, Behavior Infraction Types, Positive Note Types, and the Positive Points Cap per term
- **Staff** — add, edit, and remove staff; role (teacher/admin) set per person
- **Students** — add a student (optionally with contacts at the same time), bulk import a CSV, and browse the full roster
- **Reset & Archive** — reset term points for a mid-year reset, archive a single term's referrals to CSV without resetting anything, or Start New Year to archive Referrals/Students/ParentContacts and begin fresh

**Viewing a previously archived CSV:** the same Reset & Archive tab includes a read-only viewer for any CSV previously exported by this app (from either archive tool above) — upload the file and it displays with the same point-value coloring and readable date/time formatting used throughout the rest of the app, not a raw spreadsheet-style dump. It defaults to the columns admins actually look at (date, student, infraction, points, teacher, etc.), with a "Show all columns" toggle for the rest, plus click-to-sort columns and a quick summary (referral count, unique students, date range, total points deducted). Nothing uploaded here is ever sent to the server or written into any live sheet — it's parsed and displayed entirely in the browser.

Unsaved changes are visually flagged — a section's Save button becomes sticky and highlighted, and the tab itself shows a small dot — until it's saved, and the browser will warn before you navigate away with anything unsaved.

A search box above the tabs jumps straight to any setting by name.

---

## FERPA Compliance

This system is FERPA-compliant under these conditions:

1. ✅ School uses Google Workspace for Education (not personal Gmail)
2. ✅ Google Workspace for Education Agreement is signed with Google
3. ✅ Web App deployed with **"Anyone within [your domain]"** access — never "Anyone"
4. ✅ All teachers use school-issued Google accounts to access the form
5. ✅ Spreadsheet shared only with authorized staff
6. ✅ Google Vault enabled for audit trail (recommended)
7. ⚠️ Education law attorney review recommended before storing real student data
8. ⚠️ Software Development Agreement should be in place

---

## Gmail Daily Sending Limits

| Workspace Tier | Daily Email Limit |
|---|---|
| Education Fundamentals (free) | 100 emails/day |
| Education Standard / Plus | 1,500 emails/day |

Each referral can send up to 2 emails (contacts + teacher). A student with multiple contacts sends one email per contact as part of the same daily batch.
For high-volume schools, monitor usage or upgrade to Education Standard.

---

## Importing Your Student Roster

From **Settings → Students → Bulk Import (CSV)**, upload a CSV with columns in this order: `StudentID, FirstName, LastName, Grade`, with an optional 5th column for `MiddleName`. The first row is treated as a header and skipped. New students start at the current Term Start Points value automatically; existing StudentIDs already on the roster are skipped without error, so it's safe to re-import the same file.

**To also import a referral email contact per student** (one per row), add four more columns: `ContactType, ContactFirstName, ContactLastName, ContactEmail` (columns 6–9). Leave those four blank for any student with no contact to import yet — a full school-year roster import doesn't need every student to have a contact ready on day one. `ContactType` is matched against your configured Contact Types case-insensitively (so "parent" and "Parent" both work), but must still be a real configured type — an unrecognized one is skipped with a clear error rather than silently failing, and the student still imports either way, even if their contact row has a problem.

`ClassID`/`ClassName` are not part of the current schema — teachers are selected directly by name on the referral form, not matched by class assignment.
