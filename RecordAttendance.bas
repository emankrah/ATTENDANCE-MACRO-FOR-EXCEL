Attribute VB_Name = "Module1"
Option Explicit

'====================================================================
' ATTENDANCE AUTOMATION SYSTEM
' Reads employees from "EMPLOYEES" sheet, records attendance into
' "ATTENDANCE" sheet. Run RecordAttendance to start a session.
'====================================================================

Public Sub RecordAttendance()

    Dim wsEmp As Worksheet
    Dim wsAtt As Worksheet
    Dim continueSession As Boolean
    Dim attDate As Variant

    ' ---- Locate required worksheets ----
    On Error Resume Next
    Set wsEmp = ThisWorkbook.Worksheets("EMPLOYEES")
    Set wsAtt = ThisWorkbook.Worksheets("ATTENDANCE")
    On Error GoTo 0

    If wsEmp Is Nothing Or wsAtt Is Nothing Then
        MsgBox "Required worksheet(s) missing. Please ensure both 'EMPLOYEES' " & _
               "and 'ATTENDANCE' worksheets exist.", vbCritical, "Setup Error"
        Exit Sub
    End If

    EnsureAttendanceHeaders wsAtt

    continueSession = True

    Do While continueSession

        ' ---- Step 1: Get and validate attendance date ----
        attDate = GetAttendanceDate()
        If IsNull(attDate) Then
            ' User cancelled the date prompt
            Exit Sub
        End If

        ' ---- Step 2: Read employee list dynamically ----
        Dim lastEmpRow As Long
        lastEmpRow = wsEmp.Cells(wsEmp.Rows.Count, "A").End(xlUp).Row

        If lastEmpRow < 2 Then
            MsgBox "No employees found in the EMPLOYEES worksheet.", vbExclamation, "No Data"
            Exit Sub
        End If

        Dim totalEmployees As Long
        totalEmployees = lastEmpRow - 1  ' minus header row

        ' ---- Step 3: Determine where to start writing in ATTENDANCE ----
        Dim nextRow As Long
        Dim lastAttRow As Long
        lastAttRow = wsAtt.Cells(wsAtt.Rows.Count, "A").End(xlUp).Row

        If lastAttRow >= 2 Then
            ' Existing records present -> leave two fully blank rows before new date block
            nextRow = lastAttRow + 3
        Else
            ' Sheet only has header (or is empty) -> start right after header
            nextRow = 2
        End If

        ' ---- Step 4: Process each employee ----
        Dim i As Long
        Dim empName As String, empID As String
        Dim progressLabel As String
        Dim cancelled As Boolean
        Dim timeInResult As Variant
        Dim timeOutResult As Variant

        For i = 2 To lastEmpRow

            empName = Trim(CStr(wsEmp.Cells(i, 1).Value))
            empID = Trim(CStr(wsEmp.Cells(i, 2).Value))

            If empName = "" Then GoTo NextEmployee ' skip blank rows

            progressLabel = "Employee " & (i - 1) & " of " & totalEmployees & vbCrLf & _
                             empName & vbCrLf & empID

            ' --- Time In ---
            cancelled = False
            timeInResult = PromptTime("Time In", progressLabel, True, cancelled)
            If cancelled Then Exit Sub

            If IsBlankResult(timeInResult) Then
                ' Blank Time In => absent, no record, move to next employee
                GoTo NextEmployee
            End If

            ' --- Time Out ---
            cancelled = False
            timeOutResult = PromptTime("Time Out", progressLabel, False, cancelled)
            If cancelled Then Exit Sub

            ' --- Write record ---
            With wsAtt
                .Cells(nextRow, 1).Value = CDate(attDate)
                .Cells(nextRow, 1).NumberFormat = "dd-mm-yy"

                .Cells(nextRow, 2).Value = empName
                .Cells(nextRow, 3).Value = empID

                .Cells(nextRow, 4).Value = CDate(timeInResult)
                .Cells(nextRow, 4).NumberFormat = "h:mm AM/PM"

                If IsBlankResult(timeOutResult) Then
                    .Cells(nextRow, 5).Value = ""
                Else
                    .Cells(nextRow, 5).Value = CDate(timeOutResult)
                    .Cells(nextRow, 5).NumberFormat = "h:mm AM/PM"
                End If
            End With

            nextRow = nextRow + 1

NextEmployee:
        Next i

        ' ---- Step 5: Ask whether to continue with another date ----
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Would you like to enter attendance for another date?", _
                       vbYesNo + vbQuestion, "Continue?")

        If resp = vbNo Then
            MsgBox "Attendance recording completed successfully.", vbInformation, "Done"
            continueSession = False
        End If

    Loop

End Sub

'====================================================================
' Ensures the ATTENDANCE sheet has the correct header row if empty
'====================================================================
Private Sub EnsureAttendanceHeaders(ws As Worksheet)
    If Trim(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Cells(1, 1).Value = "DATE"
        ws.Cells(1, 2).Value = "NAME"
        ws.Cells(1, 3).Value = "STAFF ID"
        ws.Cells(1, 4).Value = "TIME IN"
        ws.Cells(1, 5).Value = "TIME OUT"
        ws.Rows(1).Font.Bold = True
    End If
End Sub

'====================================================================
' Prompts for and validates the attendance date.
' Returns a Date value, or Null if the user cancels.
'====================================================================
Private Function GetAttendanceDate() As Variant

    Dim rawInput As Variant
    Dim s As String

    Do
        rawInput = Application.InputBox("Enter Attendance Date (e.g., 20-04-26):", _
                                         "Attendance Date", Type:=2)

        ' User pressed Cancel
        If VarType(rawInput) = vbBoolean Then
            GetAttendanceDate = Null
            Exit Function
        End If

        s = Trim(CStr(rawInput))

        If s = "" Then
            MsgBox "Date cannot be blank. Please enter a valid date.", vbExclamation, "Invalid Date"
        ElseIf Not IsDate(s) Then
            MsgBox "Invalid date entered: '" & s & "'." & vbCrLf & _
                   "Please enter a valid date (e.g., 20-04-26).", vbExclamation, "Invalid Date"
        Else
            GetAttendanceDate = CDate(s)
            Exit Function
        End If
    Loop

End Function

'====================================================================
' Prompts for a time value (Time In or Time Out), validating input.
' Returns:
'   - a Date/time value if a valid time was entered
'   - "" (empty string) if the field was left blank
' Sets cancelled = True if the user pressed Cancel.
'====================================================================
Private Function PromptTime(labelText As String, employeeLabel As String, _
                             isTimeIn As Boolean, ByRef cancelled As Boolean) As Variant

    Dim rawInput As Variant
    Dim s As String
    Dim parsedTime As Date

    Do
        rawInput = Application.InputBox(labelText & ":" & vbCrLf & vbCrLf & employeeLabel, _
                                         labelText, Type:=2)

        If VarType(rawInput) = vbBoolean Then
            cancelled = True
            PromptTime = ""
            Exit Function
        End If

        s = Trim(CStr(rawInput))

        If s = "" Then
            PromptTime = ""   ' blank is allowed here; caller decides meaning
            Exit Function
        End If

        If TryParseTime(s, isTimeIn, parsedTime) Then
            PromptTime = parsedTime
            Exit Function
        Else
            MsgBox "Invalid time entered: '" & s & "'." & vbCrLf & _
                   "Please enter a valid time (e.g., 7:30, 7:30 AM, 18:30).", _
                   vbExclamation, "Invalid Time"
        End If
    Loop

End Function

'====================================================================
' Helper: checks whether a PromptTime result represents a blank entry
'====================================================================
Private Function IsBlankResult(v As Variant) As Boolean
    IsBlankResult = (VarType(v) = vbString) And (CStr(v) = "")
End Function

'====================================================================
' Parses a raw time string into a valid Date/time value.
' isTimeIn = True  -> omitted AM/PM defaults to AM (24h values pass through)
' isTimeIn = False -> omitted AM/PM defaults to PM for hours 1-11
' Returns True/False for success, and the parsed value via result.
'====================================================================
Private Function TryParseTime(raw As String, isTimeIn As Boolean, ByRef result As Date) As Boolean

    Dim s As String
    Dim upperS As String
    Dim ampm As String
    Dim parts() As String
    Dim h As Integer, m As Integer

    TryParseTime = False

    s = Trim(raw)
    upperS = UCase(s)
    ampm = ""

    If InStr(upperS, "AM") > 0 Then
        ampm = "AM"
        s = Trim(Replace(upperS, "AM", ""))
    ElseIf InStr(upperS, "PM") > 0 Then
        ampm = "PM"
        s = Trim(Replace(upperS, "PM", ""))
    End If

    If InStr(s, ":") = 0 Then Exit Function

    parts = Split(s, ":")
    If UBound(parts) <> 1 Then Exit Function

    If Not IsNumeric(Trim(parts(0))) Or Not IsNumeric(Trim(parts(1))) Then Exit Function

    h = CInt(Trim(parts(0)))
    m = CInt(Trim(parts(1)))

    If m < 0 Or m > 59 Then Exit Function

    If ampm <> "" Then
        ' 12-hour format explicitly given
        If h < 1 Or h > 12 Then Exit Function
        If ampm = "AM" Then
            If h = 12 Then h = 0
        Else ' PM
            If h <> 12 Then h = h + 12
        End If
    Else
        ' No AM/PM specified
        If h < 0 Or h > 23 Then Exit Function

        If isTimeIn Then
            ' Assume AM. Hours 0-11 already represent AM correctly;
            ' hours 13-23 are already unambiguous 24h values.
        Else
            ' Time Out: assume PM whenever hour is between 1 and 11
            If h >= 1 And h <= 11 Then
                h = h + 12
            End If
        End If
    End If

    result = TimeSerial(h, m, 0)
    TryParseTime = True

End Function

