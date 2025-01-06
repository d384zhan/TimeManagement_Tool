Attribute VB_Name = "OutputModule"
Sub PopulateCalendar()
    Dim wsCalendar As Worksheet
    Dim wsDueDates As Worksheet
    Dim dueDateRange As Range
    Dim calendarRange As Range
    Dim dueDateCell As Range
    Dim calendarCell As Range
    Dim taskName As String
    Dim dueDate As Date
    Dim calendarDate As Date
    Dim colOffset As Integer
    Dim found As Boolean
    
    ' Define worksheets
    Set wsCalendar = ThisWorkbook.Sheets("Output")
    Set wsDueDates = ThisWorkbook.Sheets("Due Dates")
    
    ' Define due dates range (Assumes headers in row 1, data starts at row 3)
    Set dueDateRange = wsDueDates.Range("D3:D" & wsDueDates.Cells(wsDueDates.Rows.Count, "D").End(xlUp).row)
    
    ' Define calendar date range (Assumes dates start in row 4S)
    Set calendarRange = wsCalendar.Range("C4:I4")
    
    ' Loop through due dates
    For Each dueDateCell In dueDateRange
        If IsDate(dueDateCell.Value) Then
            dueDate = dueDateCell.Value
            taskName = wsDueDates.Cells(dueDateCell.row, "A").Value ' Task name in column A
            
            found = False
        End If
        
        ' Find matching date in calendar
        For Each calendarCell In calendarRange
            calendarDate = calendarCell.Value
            If calendarDate = dueDate Then
                colOffset = calendarCell.Column ' Adjust to match weekday column
                found = True
                Exit For
            End If
        Next calendarCell
        
        ' If match found, insert task into corresponding time slot
        If found Then
            Dim timeSlotRow As Long
            'timeSlotRow = wsCalendar.Cells(wsCalendar.Rows.Count, colOffset).End(xlUp).row + 1
            timeSlotRow = 5
            Do While wsCalendar.Cells(timeSlotRow, colOffset) <> ""
                timeSlotRow = timeSlotRow + 1
            Loop
            
            wsCalendar.Cells(timeSlotRow, colOffset).Value = taskName
            wsCalendar.Cells(timeSlotRow, colOffset).Interior.colorIndex = 8
            
        End If
    Next dueDateCell
    
    MsgBox "Calendar updated with due dates!", vbInformation
End Sub

Sub ClearCalendar()
    Dim wsCalendar As Worksheet
    Dim startRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim lastRow As Long
    
    ' Set the worksheet and calendar range parameters
    Set wsCalendar = ThisWorkbook.Sheets("Output")
    
    startRow = 5 ' Row where the calendar content starts
    'startCol = 3 ' Column where the first day (e.g., Sunday) starts
    'endCol = wsCalendar.Cells(2, wsCalendar.Columns.Count).End(xlToLeft).Column ' Last column with weekday headers
    'lastRow = wsCalendar.Cells(wsCalendar.Rows.Count, startCol).End(xlUp).row ' Last row with data
    lastRow = 148
    
    ' Clear calendar contents
    wsCalendar.Range(wsCalendar.Cells(startRow, "C"), wsCalendar.Cells(lastRow, "I")).ClearContents
    wsCalendar.Range(wsCalendar.Cells(startRow, "C"), wsCalendar.Cells(lastRow, "I")).Interior.colorIndex = 0
    
    MsgBox "Calendar cleared successfully!", vbInformation
End Sub

Sub ShowSelectWeek()

    SelectWeekForm.Show

End Sub

Sub SwitchToDueDates()
    ' Link user to another sheet when button is clicked
    Dim ws As Worksheet

    ' Specify the target sheet name
    Set ws = ThisWorkbook.Sheets("Due Dates") ' Change "TargetSheet" to your target sheet name

    ' Activate the target sheet
    ws.Activate
End Sub
