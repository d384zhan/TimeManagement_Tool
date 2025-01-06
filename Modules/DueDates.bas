Attribute VB_Name = "DueDatesModule"
Sub ShowDueDatesForm()
    DueDatesForm.Show
End Sub

Sub DeleteAssignmentsTableRow()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowNum As Variant

    Set ws = ThisWorkbook.Worksheets("Due Dates")
    Set tbl = ws.ListObjects("AssignmentsTable")

    ' Ask the user for the row number to delete
    rowNum = InputBox("Enter the number of the assignment you would like to delete (1 deletes the first assignment, 2 deletes the second, etc):", "Delete Table Row")
    
    ' Validate user input
    If Not IsNumeric(rowNum) Or rowNum < 1 Then
        MsgBox "Invalid assignment number entered. Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    If rowNum > tbl.ListRows.Count Then
        MsgBox "The row number exceeds the number of rows in the table.", vbExclamation
        Exit Sub
    End If

    tbl.ListRows(rowNum).Delete
End Sub


Sub ClearAssignmentsTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Due Dates")
    Set tbl = ws.ListObjects("AssignmentsTable")
    
    tbl.DataBodyRange.ClearContents

End Sub

Sub UpdateAssignmentsPivotChart()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Sheets("Due Dates")
    Set tbl = ws.ListObjects("AssignmentsTable")
    
    Set pt = ThisWorkbook.Sheets("Due Dates").PivotTables("AssignmentsPivot") ' Change to your Pivot Table sheet and name
    pt.ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=tbl.Range.Address(External:=True))

    pt.RefreshTable

End Sub





