Attribute VB_Name = "SendCoursesToOutputModule"
Sub PopulateWeeklyCalendarWithCourses()
    Dim wsCourses As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim days As Variant
    Dim timeRange As Variant
    Dim courseName As String
    Dim courseType As String
    Dim startTime As Date, endTime As Date
    Dim startRow As Long, endRow As Long
    Dim dayColumn As Long
    Dim timeIncrement As Double
    Dim day As Variant
    
    ' Set worksheets
    Set wsCourses = ThisWorkbook.Sheets("Courses")
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    ' Define time increment (10 minutes = 10/1440 days)
    timeIncrement = 10 / 1440
    
    ' Loop through mandatory courses
    lastRow = wsCourses.Cells(wsCourses.Rows.Count, "C").End(xlUp).row
    For rowIndex = 7 To lastRow
        courseName = wsCourses.Cells(rowIndex, "C").Value
        If courseName <> "" Then
            days = Split(wsCourses.Cells(rowIndex, "D").Value, ", ")
            timeRange = Split(wsCourses.Cells(rowIndex, "E").Value, "-")
            startTime = TimeValue(Trim(timeRange(0)))
            endTime = TimeValue(Trim(timeRange(1)))
            
            ' Highlight each day
            For Each day In days
                dayColumn = GetDayColumn(CStr(day))
                If dayColumn <> -1 Then
                    ' Find start and end rows for the time range
                    startRow = 5 + (startTime - TimeValue("6:00 AM")) / timeIncrement
                    endRow = 5 + (endTime - TimeValue("6:00 AM")) / timeIncrement
                    
                    ' Fill in the cells
                    With wsOutput
                        .Range(.Cells(startRow, dayColumn), .Cells(endRow - 1, dayColumn)).Interior.Color = RGB(173, 216, 230) ' Light blue
                        .Cells(startRow, dayColumn).Value = courseName
                        .Cells(startRow, dayColumn).WrapText = True ' Enable text wrapping
                        .Rows(startRow & ":" & startRow).AutoFit ' Auto-adjust row height
                    End With
                End If
            Next day
        End If
    Next rowIndex
    
    ' Loop through elective courses
    lastRow = wsCourses.Cells(wsCourses.Rows.Count, "F").End(xlUp).row
    For rowIndex = 7 To lastRow
        courseName = wsCourses.Cells(rowIndex, "F").Value
        If courseName <> "" Then
            days = Split(wsCourses.Cells(rowIndex, "G").Value, ", ")
            timeRange = Split(wsCourses.Cells(rowIndex, "H").Value, "-")
            startTime = TimeValue(Trim(timeRange(0)))
            endTime = TimeValue(Trim(timeRange(1)))
            
            ' Highlight each day
            For Each day In days
                dayColumn = GetDayColumn(CStr(day))
                If dayColumn <> -1 Then
                    ' Find start and end rows for the time range
                    startRow = 5 + (startTime - TimeValue("6:00 AM")) / timeIncrement
                    endRow = 5 + (endTime - TimeValue("6:00 AM")) / timeIncrement
                    
                    ' Fill in the cells
                    With wsOutput
                        .Range(.Cells(startRow, dayColumn), .Cells(endRow - 1, dayColumn)).Interior.Color = RGB(255, 182, 193) ' Light pink
                        .Cells(startRow, dayColumn).Value = courseName
                        .Cells(startRow, dayColumn).WrapText = True ' Enable text wrapping
                        .Rows(startRow & ":" & startRow).AutoFit ' Auto-adjust row height
                    End With
                End If
            Next day
        End If
    Next rowIndex
    
    MsgBox "Weekly calendar updated successfully!"
End Sub

Function GetDayColumn(day As String) As Long
    Select Case day
        Case "Sun": GetDayColumn = 3
        Case "Mon": GetDayColumn = 4
        Case "Tue": GetDayColumn = 5
        Case "Wed": GetDayColumn = 6
        Case "Thu": GetDayColumn = 7
        Case "Fri": GetDayColumn = 8
        Case "Sat": GetDayColumn = 9
        Case Else: GetDayColumn = -1
    End Select
End Function

