' ── SendParentNotifications ────────────────────────────────────────────
' Reads parent_notifications.csv and sends one email per row "from" a
' shared mailbox, using SENT AS permission (not Full Access) — set via
' SentOnBehalfOfName, not SendUsingAccount. See notes below on what that
' means in practice.
'
' PERMISSION MODEL: this mailbox was set up with Send As permission only
' (not Full Access), so it does NOT appear as an account in this Outlook
' profile — SendUsingAccount would not find it. SentOnBehalfOfName is the
' correct property for Send As: with genuine Send As permission, Exchange
' sends the message as truly FROM the shared mailbox, with no "on behalf
' of" wording shown to the recipient (that wording only appears with the
' separate "Send on Behalf" permission, which this isn't).
'
' IMPORTANT — sent items land in YOUR mailbox, not the shared one: since
' this isn't Full Access, Outlook never actually opens the shared mailbox
' as a second mailbox — it just stamps the From address on a message
' that's really being sent through your own account. That means sent
' copies are saved to YOUR OWN Sent Items, not the shared mailbox's. If
' you ever need to check what's been sent, look there.
'
' A permission problem here may not surface as an immediate VBA error —
' if Send As isn't actually granted correctly, Exchange can accept the
' message and then bounce it back as an NDR (non-delivery report) email
' later, which this macro has no way to detect in real time. If Sent
' counts look right but a parent says they never got anything, check
' your own Inbox for a bounce-back before assuming it's a different
' problem.
'
' NOTE ON "MARKED AS SENT": BehaviorTracker already marks every exported
' referral as notified the moment the CSV is downloaded — that happens
' in the app, not here. This macro has nothing to report back and
' doesn't touch the spreadsheet at all. If a send in here is later found
' to have failed, go set that referral's Parent Notification back to
' "No" in the Report page so it gets picked up again next time.
'
' SETUP (one-time):
'   1. Open Excel, press Alt+F11 to open the VBA editor.
'   2. Insert > Module, paste this entire file's contents in.
'   3. Update SHARED_MAILBOX_SMTP below.
'   4. Tools > References > check "Microsoft Outlook XX.0 Object Library"
'      if it isn't already checked (lets this talk to Outlook).
'   5. Save the workbook as .xlsm (macro-enabled) — a plain .xlsx will
'      silently discard this code on save.
'   6. Run SendParentNotifications() via F5, or assign it to a button
'      (Developer tab > Insert > Button > assign macro) for one-click use.
'
' REQUIRES: Windows, Classic Outlook (not New Outlook), signed in as the
' account that has Send As permission on the shared mailbox.
' ─────────────────────────────────────────────────────────────────────

Const SHARED_MAILBOX_SMTP As String = "behavior.tracker@calloway.kyschools.us" ' TODO: set this
Const SUBJECT_PREFIX As String = "Daily Behavior Summary"

Sub SendParentNotifications()

    Dim csvPath As String
    Dim OutApp As Object, OutMail As Object
    Dim wbCsv As Workbook, wsCsv As Worksheet
    Dim lastRow As Long, i As Long
    Dim sentCount As Long, failCount As Long
    Dim failLog As String
    Dim colDate As Integer, colParentName As Integer, colParentEmail As Integer
    Dim colStudent As Integer, colBody As Integer
    Dim c As Integer

    ' ── 1. Pick the CSV ──────────────────────────────────────────────
    csvPath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Select today's parent_notifications.csv")
    If csvPath = "False" Then Exit Sub ' user cancelled

    Set wbCsv = Workbooks.Open(csvPath)
    Set wsCsv = wbCsv.Sheets(1)
    lastRow = wsCsv.Cells(wsCsv.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data rows found in that CSV.", vbExclamation
        wbCsv.Close SaveChanges:=False
        Exit Sub
    End If

    ' ── 2. Find columns by header name, not fixed position ──────────
    ' (safer than assuming column order — matches whatever order
    ' downloadNotifCsv() in Report.html happens to export in. The CSV
    ' also has a ReferralIDs column, which this macro doesn't need —
    ' BehaviorTracker already marked everything notified at export time.)
    For c = 1 To wsCsv.Cells(1, wsCsv.Columns.Count).End(xlToLeft).Column
        Select Case Trim(wsCsv.Cells(1, c).Value)
            Case "Date":            colDate = c
            Case "ParentName":      colParentName = c
            Case "ParentEmail":     colParentEmail = c
            Case "StudentDisplay":  colStudent = c
            Case "Body":            colBody = c
        End Select
    Next c

    If colParentEmail = 0 Or colBody = 0 Or colStudent = 0 Or colDate = 0 Then
        MsgBox "CSV is missing an expected column (Date, ParentEmail, StudentDisplay, or Body). " & _
               "Check that this is the right file.", vbCritical
        wbCsv.Close SaveChanges:=False
        Exit Sub
    End If

    ' ── 3. Confirm before sending anything ───────────────────────────
    If MsgBox("This will send " & (lastRow - 1) & " email(s) as " & SHARED_MAILBOX_SMTP & _
              " (via Send As). These referrals are already marked as notified in " & _
              "BehaviorTracker — only proceed if you're ready to actually send them now. " & _
              "Continue?", vbYesNo + vbQuestion, "Confirm Send") <> vbYes Then
        wbCsv.Close SaveChanges:=False
        Exit Sub
    End If

    ' ── 4. Connect to Outlook ─────────────────────────────────────
    ' No account lookup needed here, unlike the Full-Access version of
    ' this macro — Send As doesn't add the mailbox to this profile, so
    ' there's nothing to find. SentOnBehalfOfName (set per-message below)
    ' is what does the actual work.
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")

    ' ── 5. Send one email per row ─────────────────────────────────
    sentCount = 0: failCount = 0: failLog = ""

    For i = 2 To lastRow
        On Error GoTo RowFailed

        Dim rowDate As String, rowParentName As String, rowParentEmail As String
        Dim rowStudent As String, rowBody As String

        rowDate = Trim(wsCsv.Cells(i, colDate).Value)
        rowParentName = Trim(wsCsv.Cells(i, colParentName).Value)
        rowParentEmail = Trim(wsCsv.Cells(i, colParentEmail).Value)
        rowStudent = Trim(wsCsv.Cells(i, colStudent).Value)
        rowBody = wsCsv.Cells(i, colBody).Value

        If rowParentEmail <> "" Then
            Set OutMail = OutApp.CreateItem(0) ' 0 = olMailItem
            With OutMail
                .SentOnBehalfOfName = SHARED_MAILBOX_SMTP
                .To = rowParentEmail
                .Subject = SUBJECT_PREFIX & " — " & rowStudent & " — " & rowDate
                .Body = rowBody
                .Send
            End With
            sentCount = sentCount + 1
            Set OutMail = Nothing
        End If

        GoTo NextRow

RowFailed:
        failCount = failCount + 1
        failLog = failLog & vbCrLf & "Row " & i & " (" & rowParentEmail & "): " & Err.Description
        Err.Clear
        Resume NextRow

NextRow:
        On Error GoTo 0
    Next i

    wbCsv.Close SaveChanges:=False

    ' ── 6. Report results ───────────────────────────────────────────
    Dim summary As String
    summary = "Sent: " & sentCount & vbCrLf & "Failed: " & failCount
    If failCount > 0 Then
        summary = summary & vbCrLf & vbCrLf & _
            "Any failed rows above are already marked 'notified' in BehaviorTracker " & _
            "(that happened at export time) — go to the Report page and set Parent " & _
            "Notification back to 'No' for each one so it's picked up again next time." & _
            vbCrLf & failLog
    End If
    summary = summary & vbCrLf & vbCrLf & _
        "Reminder: sent copies are in YOUR Sent Items, not the shared mailbox's " & _
        "(that's expected with Send As, not a bug)."
    MsgBox summary, IIf(failCount > 0, vbExclamation, vbInformation), "Done"

End Sub
