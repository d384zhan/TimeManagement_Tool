VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TaskFinish 
   Caption         =   "UserForm1"
   ClientHeight    =   3437
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   5760
   OleObjectBlob   =   "TaskFinishForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TaskFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Reference the "Tasks" sheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    ' Clear cboxTask to avoid duplicates
    cboxTask.Clear

    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row

    ' Populate cboxTask with tasks from column B starting from row 4
    For i = 4 To lastRow
        If ws.Cells(i, 2).Value <> "" Then
            cboxTask.AddItem ws.Cells(i, 2).Value
        End If
    Next i
End Sub

Private Sub SubmitBtn_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim selectedTask As String
    Dim rowToDelete As Long

    ' Reference the "Tasks" sheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    ' Ensure a task is selected
    If cboxTask.Value = "" Then
        MsgBox "Please select a task to delete.", vbExclamation
        Exit Sub
    End If

    ' Get the selected task
    selectedTask = cboxTask.Value

    ' Find the row of the selected task
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    rowToDelete = 0
    For i = 4 To lastRow
        If ws.Cells(i, 2).Value = selectedTask Then
            rowToDelete = i
            Exit For
        End If
    Next i

    ' If the task is found, delete the row and shift data up
    If rowToDelete > 0 Then
        ws.Rows(rowToDelete).Delete Shift:=xlUp
        MsgBox "Task deleted successfully!", vbInformation
    Else
        MsgBox "Task not found. Please try again.", vbExclamation
    End If

End Sub
