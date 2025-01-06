Attribute VB_Name = "CoursesModule"
Sub ShowStudentInfoForm()
    StudentInfoForm.Show
End Sub

Sub ClearCoursesData()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Courses")
    On Error GoTo 0

    ' Check if the sheet exists
    If ws Is Nothing Then
        MsgBox "The sheet 'Courses' does not exist.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Clear C4 and C5
    ws.Range("C4:C5").ClearContents

    ' Find the last row with data in columns C to H
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' Ensure there is data to clear from row 7 onward
    If lastRow >= 7 Then
        ws.Range("C7:H" & lastRow).ClearContents
    End If

End Sub

Sub GetCourseDetails()
    Dim wsCourses As Worksheet
    Dim wsInfo As Worksheet
    Dim programName As String
    Dim term As String
    Dim col As Long, row As Long
    Dim courseName As String
    Dim startTime As String, endTime As String, timeRange As String
    Dim electiveCount As Integer
    Dim targetCol As Long
    Dim targetRow As Long
    Dim lastRow As Long
    Dim writeRow As Long
    Dim electiveWriteRow As Long
    Dim days As String
    Dim daySelections As String

    ' Set the worksheets
    Set wsCourses = ThisWorkbook.Sheets("Courses")
    Set wsInfo = ThisWorkbook.Sheets("Courses Info Sheet")
    
    ' Get the program name and term from the "Courses" sheet
    programName = wsCourses.Range("C4").Value
    term = wsCourses.Range("C5").Value
    
    ' Find the column in "Courses Info Sheet" with the program name
    targetCol = 0
    For col = 1 To wsInfo.Cells(1, wsInfo.Columns.Count).End(xlToLeft).Column
        If wsInfo.Cells(1, col).Value = programName Then
            targetCol = col
            Exit For
        End If
    Next col
    
    If targetCol = 0 Then
        MsgBox "Program not found in 'Courses Info Sheet'.", vbExclamation
        Exit Sub
    End If
    
    ' Find the row with the specified term in the target column
    targetRow = 0
    lastRow = wsInfo.Cells(wsInfo.Rows.Count, targetCol).End(xlUp).row
    For row = 1 To lastRow
        If wsInfo.Cells(row, targetCol).Value = term Then
            targetRow = row
            Exit For
        End If
    Next row
    
    If targetRow = 0 Then
        MsgBox "Term not found in 'Courses Info Sheet'.", vbExclamation
        Exit Sub
    End If
    
    ' Start writing data to the "Courses" sheet
    writeRow = 7 ' Start writing course names in column C from row 7
    electiveWriteRow = 7 ' Start writing elective names in column F from row 7
    
    ' Process the courses below the term
    row = targetRow + 1
    While IsNumeric(wsInfo.Cells(row, targetCol).Value) = False
        courseName = wsInfo.Cells(row, targetCol).Value
        
        ' Ask for start and end time
        timeRange = InputBox("Enter the start and end times for " & courseName & " in 24-hour format (e.g., 13:00-14:30):", "Course Time")
        
        ' Validate the input format
        If Len(timeRange) < 11 Or InStr(timeRange, "-") = 0 Then
            MsgBox "Invalid time format. Please enter times in 24-hour format (e.g., 13:00-14:30).", vbExclamation
            Exit Sub
        End If
        
        startTime = Left(timeRange, 5)
        endTime = Mid(timeRange, 7)
        
        ' Ask for class days
        daySelections = GetClassDays(courseName)
        
        ' Write the course details in "Courses" sheet
        wsCourses.Cells(writeRow, "C").Value = courseName
        wsCourses.Cells(writeRow, "D").Value = daySelections
        wsCourses.Cells(writeRow, "E").Value = timeRange
        
        writeRow = writeRow + 1
        row = row + 1
    Wend
    
    ' Handle electives
    electiveCount = wsInfo.Cells(row, targetCol).Value
    If electiveCount > 0 Then
        For i = 1 To electiveCount
            courseName = InputBox("Enter the name of elective course " & i & ":", "Elective Course Name")
            timeRange = InputBox("Enter the start and end times for " & courseName & " in 24-hour format (e.g., 13:00-14:30):", "Elective Time")
            
            ' Validate the input format
            If Len(timeRange) < 11 Or InStr(timeRange, "-") = 0 Then
                MsgBox "Invalid time format. Please enter times in 24-hour format (e.g., 13:00-14:30).", vbExclamation
                Exit Sub
            End If
            
            startTime = Left(timeRange, 5)
            endTime = Mid(timeRange, 7)
            
            daySelections = GetClassDays(courseName)
            
            ' Write the elective details in "Courses" sheet
            wsCourses.Cells(electiveWriteRow, "F").Value = courseName
            wsCourses.Cells(electiveWriteRow, "G").Value = daySelections
            wsCourses.Cells(electiveWriteRow, "H").Value = timeRange
            
            electiveWriteRow = electiveWriteRow + 1
        Next i
    End If
    
    ' Auto-fit columns to ensure all data is visible
    wsCourses.Columns("C:H").AutoFit
    
    MsgBox "Course details have been collected successfully!", vbInformation
End Sub

Function GetClassDays(courseName As String) As String
    Dim daySelections As String
    Dim days(6) As String
    Dim dayResponses(6) As Variant
    Dim i As Integer
    
    days(0) = "Monday"
    days(1) = "Tuesday"
    days(2) = "Wednesday"
    days(3) = "Thursday"
    days(4) = "Friday"
    days(5) = "Saturday"
    days(6) = "Sunday"
    
    daySelections = ""
    
    ' Ask user to check off days
    For i = 0 To 6
        dayResponses(i) = MsgBox("Is " & days(i) & " a class day for " & courseName & "?", vbYesNo, "Class Days")
        If dayResponses(i) = vbYes Then
            daySelections = daySelections & Left(days(i), 3) & ", "
        End If
    Next i
    
    ' Remove trailing comma and space
    If Len(daySelections) > 0 Then
        daySelections = Left(daySelections, Len(daySelections) - 2)
    End If
    
    GetClassDays = daySelections
End Function

