VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StudentInfoForm 
   Caption         =   "Enter Info"
   ClientHeight    =   5243
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   6400
   OleObjectBlob   =   "StudentInfoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StudentInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Add items to programBox
    With programBox
        .AddItem "Architectural Engineering"
        .AddItem "Architecture"
        .AddItem "Biomedical Engineering"
        .AddItem "Chemical Engineering"
        .AddItem "Civil Engineering"
        .AddItem "Computer Engineering"
        .AddItem "Electrical Engineering"
        .AddItem "Environmental Engineering"
        .AddItem "Geological Engineering"
        .AddItem "Management Engineering"
        .AddItem "Mechanical Engineering"
        .AddItem "Mechatronics Engineering"
        .AddItem "Nanotechnology Engineering"
        .AddItem "Software Engineering"
        .AddItem "Systems Design Engineering"
    End With

    ' Add items to termBox
    With termBox
        .AddItem "1A"
        .AddItem "1B"
        .AddItem "2A"
        .AddItem "2B"
        .AddItem "3A"
        .AddItem "3B"
        .AddItem "4A"
        .AddItem "4B"
    End With
End Sub

Private Sub programInfoBtn_Click()
    Dim programValid As Boolean
    Dim termValid As Boolean
    Dim i As Integer

    ' Validate programBox input
    programValid = False
    For i = 0 To programBox.ListCount - 1
        If programBox.Value = programBox.List(i) Then
            programValid = True
            Exit For
        End If
    Next i

    If Not programValid Then
        MsgBox "Please select a valid program from the dropdown list.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Validate termBox input
    termValid = False
    For i = 0 To termBox.ListCount - 1
        If termBox.Value = termBox.List(i) Then
            termValid = True
            Exit For
        End If
    Next i

    If Not termValid Then
        MsgBox "Please select a valid term from the dropdown list.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Write the values to the "Courses" sheet
    On Error Resume Next
    With ThisWorkbook.Sheets("Courses")
        .Range("C4").Value = programBox.Value ' Write program to C4
        .Range("C5").Value = termBox.Value    ' Write term to C5
    End With
    On Error GoTo 0


    ' Close the form
    Unload Me
End Sub

