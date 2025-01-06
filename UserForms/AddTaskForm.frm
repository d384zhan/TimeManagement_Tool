VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddTaskForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4683
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   6060
   OleObjectBlob   =   "AddTaskForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddTaskForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Clear the existing items in the combo boxes
    listType.Clear
    listTime.Clear

    ' Populate listType and listTime with predefined options
    With listType
        .AddItem "Chore"
        .AddItem "Health"
        .AddItem "Social"
        .AddItem "Meeting"
        .AddItem "Other"
    End With

    With listTime
        .AddItem "No Preference"
        .AddItem "Early Morning"
        .AddItem "Morning"
        .AddItem "Afternoon"
        .AddItem "Evening"
        .AddItem "Night"
    End With
End Sub

Private Sub SubmitBtn_Click()
    Dim ws As Worksheet
    Dim emptyRow As Long
    Dim taskDate As Date
    Dim dayVal As Integer, monthVal As Integer, yearVal As Integer
    Dim estimateVal As Integer

    ' Reference the "Tasks" sheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    ' Ensure all fields have values
    If txtItem.Value = "" Then
        MsgBox "Please enter a task item.", vbExclamation
        txtItem.SetFocus
        Exit Sub
    End If

    If txtDay.Value = "" Or txtMonth.Value = "" Or txtYear.Value = "" Then
        MsgBox "Please complete the date fields.", vbExclamation
        txtDay.SetFocus
        Exit Sub
    End If

    If listType.Value = "" Then
        MsgBox "Please select a task type.", vbExclamation
        listType.SetFocus
        Exit Sub
    End If

    If txtEstimate.Value = "" Then
        MsgBox "Please enter an estimated time.", vbExclamation
        txtEstimate.SetFocus
        Exit Sub
    End If

    If listTime.Value = "" Then
        MsgBox "Please select a preferred time.", vbExclamation
        listTime.SetFocus
        Exit Sub
    End If

    ' Validate and construct the date
    On Error Resume Next
    dayVal = CInt(txtDay.Value)
    monthVal = CInt(txtMonth.Value)
    yearVal = CInt(txtYear.Value)
    On Error GoTo 0

    If monthVal < 1 Or monthVal > 12 Then
        MsgBox "Invalid month! Please enter a value between 1 and 12.", vbExclamation
        txtMonth.SetFocus
        Exit Sub
    End If

    If dayVal < 1 Or dayVal > day(DateSerial(yearVal, monthVal + 1, 0)) Then
        MsgBox "Invalid day! Please enter a valid day for the selected month.", vbExclamation
        txtDay.SetFocus
        Exit Sub
    End If

    If yearVal <= 1900 Then
        MsgBox "Invalid year! Please enter a year greater than 1900.", vbExclamation
        txtYear.SetFocus
        Exit Sub
    End If

    ' Construct the date
    On Error Resume Next
    taskDate = DateSerial(yearVal, monthVal, dayVal)
    If Err.Number <> 0 Then
        MsgBox "Invalid date! Please check your inputs.", vbExclamation
        txtDay.SetFocus
        Exit Sub
    End If
    On Error GoTo 0

    ' Validate the estimate input
    On Error Resume Next
    estimateVal = CInt(txtEstimate.Value)
    On Error GoTo 0

    If estimateVal Mod 10 <> 0 Then
        MsgBox "Estimate must be a multiple of 10. Please re-enter.", vbExclamation
        txtEstimate.SetFocus
        Exit Sub
    End If

    ' Find the first empty row starting from row 4
    emptyRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row + 1
    If emptyRow < 4 Then emptyRow = 4

    ' Additional check to handle residual data below row 4
    If ws.Cells(emptyRow, 2).Value <> "" Then
        emptyRow = ws.Cells(4, 2).End(xlDown).row + 1
        If IsEmpty(ws.Cells(emptyRow, 2).Value) = False Then emptyRow = emptyRow + 1
    End If

    ' Write the data into the sheet
    With ws
        .Cells(emptyRow, 2).Value = txtItem.Value
        .Cells(emptyRow, 3).Value = estimateVal
        .Cells(emptyRow, 4).Value = listType.Value
        .Cells(emptyRow, 5).Value = Format(taskDate, "MM/DD/YYYY")
        .Cells(emptyRow, 6).Value = listTime.Value
    End With

    MsgBox "Task added successfully!", vbInformation

    ' Optionally, clear the form inputs after submission
    txtItem.Value = ""
    txtDay.Value = ""
    txtMonth.Value = ""
    txtYear.Value = ""
    txtEstimate.Value = ""
    listType.Value = ""
    listTime.Value = ""
    txtItem.SetFocus
End Sub

