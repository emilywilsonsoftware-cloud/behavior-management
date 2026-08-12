' ── SendParentNotifications ────────────────────────────────────────────
' Reads parent_notifications.csv and sends one email per row from a
' shared mailbox, with a per-recipient subject that includes the date.
'
' SETUP (one-time):
'   1. Open Excel, press Alt+F11 to open the VBA editor.
'   2. Insert > Module, paste this entire file's contents in.
'   3. Update the constants just below (shared mailbox address, CSV path
'      if you want it fixed instead of prompted).
'   4. Tools > References > check "Microsoft Outlook XX.0 Object Library"
'      if it isn't already checked (lets this talk to Outlook).
'   5. Save the workbook as .xlsm (macro-enabled) — a plain .xlsx will
'      silently discard this code on save.
'   6. Run SendParentNotifications() via F5, or assign it to a button
'      (Developer tab > Insert > Button > assign macro) for one-click use.
'
' REQUIRES: Classic Outlook (not New Outlook), already running/signed in
' with the shared mailbox visible as an account. See the note in chat
' about auto-mapping / Send As permission if the account isn't showing up.
' ─────────────────────────────────────────────────────────────────────

Const SHARED_MAILBOX_SMTP As String = "behavior.tracker@calloway.kyschools.us" ' TODO: set this
Const SUBJECT_PREFIX As String = "Daily Behavior Summary"

Sub SendParentNotifications()

    Dim csvPath As String
    Dim OutApp As Object, OutMail As Object, SendAccount As Object
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
    ' downloadNotifCsv() in Report.html happens to export in)
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
    If MsgBox("This will send " & (lastRow - 1) & " email(s) from " & SHARED_MAILBOX_SMTP & _
              ". Continue?", vbYesNo + vbQuestion, "Confirm Send") <> vbYes Then
        wbCsv.Close SaveChanges:=False
        Exit Sub
    End If

    ' ── 4. Connect to Outlook and find the shared-mailbox account ───
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")

    Set SendAccount = FindAccount(OutApp, SHARED_MAILBOX_SMTP)
    If SendAccount Is Nothing Then
        MsgBox "Couldn't find " & SHARED_MAILBOX_SMTP & " as an account in this Outlook profile." & _
               vbCrLf & vbCrLf & _
               "Check with IT that the shared mailbox has been added with Full Access " & _
               "(which normally auto-adds it as a selectable account), or that it's been " & _
               "manually added under File > Account Settings.", vbCritical
        wbCsv.Close SaveChanges:=False
        Exit Sub
    End If

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
                .SendUsingAccount = SendAccount
                .To = rowParentEmail
                .Subject = SUBJECT_PREFIX & " - " & rowDate
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
    If failCount > 0 Then summary = summary & vbCrLf & failLog
    MsgBox summary, IIf(failCount > 0, vbExclamation, vbInformation), "Done"

End Sub

' Finds an Outlook Account object by SMTP address (case-insensitive).
Function FindAccount(OutApp As Object, smtpAddress As String) As Object
    Dim acct As Object
    For Each acct In OutApp.Session.Accounts
        If LCase(acct.SmtpAddress) = LCase(smtpAddress) Then
            Set FindAccount = acct
            Exit Function
        End If
    Next acct
    Set FindAccount = Nothing
End Function
