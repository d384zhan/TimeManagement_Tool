VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DueDatesForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4776
   ClientLeft      =   -40
   ClientTop       =   -140
   ClientWidth     =   6920
   OleObjectBlob   =   "DueDatesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DueDatesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet
    Dim courseCell As Range
    Dim lastRow As Long
    
    ' Set reference to the "Courses" sheet
    Set ws = ThisWorkbook.Sheets("Courses")
    
    ' Find the last row with data in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' Loop through column C starting from row 7
    For Each courseCell In ws.Range("C7:C" & lastRow)
        If Trim(courseCell.Value) <> "" Then ' Check if the cell is not empty
            Me.CBCourse.AddItem courseCell.Value
        End If
    Next courseCell
    
    With Me.CBType
        .AddItem "Project"
        .AddItem "Test"
        .AddItem "Quiz"
        .AddItem "Exam"
        .AddItem "Assignment"
    End With
    
    With Me.CBStatus
        .AddItem "NOT STARTED"
        .AddItem "IN PROGRESS"
    End With
    
    With Me.CBPriority
        .AddItem "HIGH"
        .AddItem "MEDIUM"
        .AddItem "LOW"
    End With
End Sub
Private Sub SubmitButton_Click()
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Due Dates")
    
    'Submission for Assignment Information
    intRow = 3
    If (nametxt.Value <> "") Then
        If (CBCourse.Value <> "" And CBType.Value <> "") Then
            Do While (ws.Cells(intRow, "A") <> "")
                intRow = intRow + 1
            Loop
            
            ws.Cells(intRow, "A") = nametxt.Value
            ws.Cells(intRow, "B") = CBCourse.Value
            ws.Cells(intRow, "C") = CBType.Value
        Else
            MsgBox "Please enter a course and assignment type."
            nametxt.Value = ""
            CBCourse.Value = ""
            CBType.Value = ""
     
            daytxt.Value = ""
            monthtxt.Value = ""
            yeartxt.Value = ""
            CBPriority.Value = ""
            CBStatus.Value = ""
            DueDatesForm.Hide
            Exit Sub
        End If
    Else
        MsgBox "Please enter a name for the assignment."
    End If
    
    'Submission for Assignment Status
    intRow = 3
    
    If (CBStatus.Value <> "" And CBPriority.Value <> "") Then
        If (daytxt.Value <> "" And yeartxt.Value <> "" And monthtxt <> "" And IsNumeric(daytxt.Value) And IsNumeric(yeartxt.Value) And IsNumeric(monthtxt.Value)) Then
            Do While (ws.Cells(intRow, "D") <> "")
                intRow = intRow + 1
            Loop
            
            ws.Cells(intRow, "D") = yeartxt.Value + "-" + monthtxt.Value + "-" + daytxt.Value
            ws.Cells(intRow, "D").NumberFormat = "yyyy-mm-dd;@"
            ws.Cells(intRow, "E") = CBStatus.Value
            ws.Cells(intRow, "F") = CBPriority.Value
        Else
            MsgBox "Please enter a valid numeric day, month, and year."
            nametxt.Value = ""
            CBCourse.Value = ""
            CBType.Value = ""
     
            daytxt.Value = ""
            monthtxt.Value = ""
            yeartxt.Value = ""
            CBPriority.Value = ""
            CBStatus.Value = ""
            DueDatesForm.Hide
            Exit Sub
        End If
    Else
        MsgBox "Please enter a status and priority"
        nametxt.Value = ""
        CBCourse.Value = ""
        CBType.Value = ""
                 
        daytxt.Value = ""
        monthtxt.Value = ""
        yeartxt.Value = ""
        CBPriority.Value = ""
        CBStatus.Value = ""
        DueDatesForm.Hide
        Exit Sub
    End If
        
    nametxt.Value = ""
    CBCourse.Value = ""
    CBType.Value = ""
     
    daytxt.Value = ""
    monthtxt.Value = ""
    yeartxt.Value = ""
    CBPriority.Value = ""
    CBStatus.Value = ""
    DueDatesForm.Hide
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub





