VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectWeekForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2863
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   4280
   OleObjectBlob   =   "SelectWeekForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectWeekForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim currentDate As Date
    Dim sundayDate As Date
    Dim i As Integer

    ' Clear listWeeks to avoid duplication
    listWeeks.Clear

    ' Get today's date
    currentDate = Date

    ' Calculate the Sunday of the current week
    sundayDate = currentDate - Weekday(currentDate, vbSunday) + 1

    ' Populate listWeeks with the 50 closest Sundays
    For i = 0 To 49
        listWeeks.AddItem Format(sundayDate + (i * 7), "mm/dd/yyyy")
    Next i
End Sub

Private Sub SubmitBtn_Click()
    Dim ws As Worksheet
    Dim selectedSunday As Date
    Dim i As Integer

    ' Ensure a date is selected
    If listWeeks.ListIndex = -1 Then
        MsgBox "Please select a Sunday from the list.", vbExclamation
        Exit Sub
    End If

    ' Get the selected Sunday date
    selectedSunday = CDate(listWeeks.Value)

    ' Reference the "Output" sheet
    Set ws = ThisWorkbook.Sheets("Output")

    ' Write the week dates (Sunday to Saturday) in the row starting at C4
    For i = 0 To 6
        ws.Cells(4, 3 + i).Value = Format(selectedSunday + i, "mm/dd/yyyy")
    Next i

    MsgBox "This week has been updated!", vbInformation
    
    Call PopulateWeeklyCalendarWithCourses
    Call ScheduleTasks
    
End Sub

