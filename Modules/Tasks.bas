Attribute VB_Name = "TasksModule"
Sub ShowAddTasks()

    AddTaskForm.Show

End Sub

Sub ScheduleTasks()
    Dim wsTasks As Worksheet, wsOutput As Worksheet
    Dim taskRow As Long, lastTaskRow As Long
    Dim outputCol As Integer, outputRow As Long
    Dim taskDate As Date, taskDescription As String
    Dim estimateMinutes As Long, preferredTime As String
    Dim dayStartRow As Long, dayEndRow As Long, blockDuration As Long
    Dim startRow As Long, timeSlotsNeeded As Long, slotFound As Boolean
    Dim i As Long, j As Long, dayColumn As Integer
    Dim dateMatch As Boolean
    Dim outputRange As Range
    Dim taskColor As Long
    Dim colorIndex As Long

    ' Define worksheets
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsOutput = ThisWorkbook.Sheets("Output")

    ' Time block definitions
    dayStartRow = 5 ' First time block (6:00 AM)
    dayEndRow = 148 ' Last time block (5:50 AM)
    blockDuration = 10 ' Each block represents 10 minutes

    ' Define an array of colors for differentiation
    Dim taskColors As Variant
    taskColors = Array(RGB(135, 206, 250), RGB(255, 182, 193), RGB(144, 238, 144), _
                       RGB(255, 255, 102), RGB(221, 160, 221), RGB(240, 128, 128), _
                       RGB(173, 216, 230), RGB(250, 128, 114), RGB(152, 251, 152), _
                       RGB(255, 228, 181)) ' Light blue, pink, green, etc.

    colorIndex = 0 ' Start with the first color in the array

    ' Get the last row of tasks in the "Tasks" sheet
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, 2).End(xlUp).row

    ' Loop through tasks in the "Tasks" sheet starting from row 4
    For taskRow = 4 To lastTaskRow
        taskDate = wsTasks.Cells(taskRow, 5).Value ' Column E (Task Date)
        taskDescription = wsTasks.Cells(taskRow, 2).Value ' Column B (Task Description)
        estimateMinutes = wsTasks.Cells(taskRow, 3).Value ' Column C (Estimate in minutes)
        preferredTime = wsTasks.Cells(taskRow, 6).Value ' Column F (Preferred Time)

        ' Calculate number of 10-minute blocks needed
        timeSlotsNeeded = estimateMinutes / blockDuration

        ' Find the corresponding column in "Output" where the date matches
        dateMatch = False
        For outputCol = 3 To 9 ' Columns C to I
            If wsOutput.Cells(4, outputCol).Value = taskDate Then
                dayColumn = outputCol
                dateMatch = True
                Exit For
            End If
        Next outputCol

        ' If no matching date is found, skip the task
        If Not dateMatch Then
            MsgBox "Task '" & taskDescription & "' on " & Format(taskDate, "mm/dd/yyyy") & " cannot be scheduled: Date not found in Output.", vbExclamation
            GoTo NextTask
        End If

        ' Determine the earliest valid start row based on preferred time
        Select Case preferredTime
            Case "Early Morning": startRow = dayStartRow + ((7 - 6) * 6)
            Case "Morning": startRow = dayStartRow + ((10 - 6) * 6)
            Case "Afternoon": startRow = dayStartRow + ((13 - 6) * 6)
            Case "Evening": startRow = dayStartRow + ((16 - 6) * 6)
            Case "Night": startRow = dayStartRow + ((20 - 6) * 6)
            Case Else: startRow = dayStartRow + ((9 - 6) * 6) ' Default to No Preference
        End Select

        ' Look for available blocks in the Output sheet
        slotFound = False
        For outputRow = startRow To dayEndRow - timeSlotsNeeded + 1
            ' Check if all required blocks are free
            slotFound = True
            For j = 0 To timeSlotsNeeded - 1
                Set outputRange = wsOutput.Cells(outputRow + j, dayColumn)
                If Not IsEmpty(outputRange.Value) Or outputRange.Interior.colorIndex <> xlNone Then
                    slotFound = False
                    Exit For
                End If
            Next j

            ' If a valid slot is found, schedule the task
            If slotFound Then
                wsOutput.Cells(outputRow, dayColumn).Value = taskDescription ' Write task description
                taskColor = taskColors(colorIndex Mod (UBound(taskColors) + 1)) ' Cycle through colors
                For j = 0 To timeSlotsNeeded - 1
                    Set outputRange = wsOutput.Cells(outputRow + j, dayColumn)
                    outputRange.Interior.Color = taskColor ' Highlight block with unique color
                Next j
                colorIndex = colorIndex + 1 ' Move to the next color
                Exit For
            End If
        Next outputRow

        ' If no slot is found, output an error message
        If Not slotFound Then
            MsgBox "Task '" & taskDescription & "' on " & Format(taskDate, "mm/dd/yyyy") & " cannot be scheduled: No available time slots.", vbExclamation
        End If

NextTask:
    Next taskRow

    MsgBox "Task scheduling complete.", vbInformation
End Sub

Sub showTaskFinished()

    TaskFinish.Show

End Sub
