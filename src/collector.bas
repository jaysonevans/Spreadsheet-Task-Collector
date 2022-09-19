' Spreadsheet Task Collector
' This program parses through spreadsheets and retrieves tasks due in a specified
' amount of time.
'
' 
' Copyright (c) 2022 Jayson Evans
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the "Software"), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
' to whom the Software is furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all copies
' or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.

' Control the execution of other subroutines and functions
Public Sub RefreshUpcoming()

Call PrepareUpcoming
Call LookForDue

Worksheets("Upcoming").Activate

Call SortByDueDate

End Sub

' Prepare the sheet which holds upcoming task
Sub PrepareUpcoming()

With Worksheets("Upcoming")
    Cells.ClearContents
    Cells.ClearFormats
    
    ' These are categories typically found on a To-Do list
    ' Note: you may need to change the positions of interest this
    ' program uses to figure things out
    Range("A1") = "Type"
    Range("B1") = "Task"
    Range("C1") = "Due"
    Range("D1") = "Completed"
    Range("E1") = "Time (min)"
    Range("F1") = "Est (min)"
    
    '  Heading is made bold
    Range("A1", "F1").Font.Bold = True
End With

End Sub

' Look for any imminent due dates
Sub LookForDue()

' Used for color coding tasks
' nameColor.Add "YOUR_SHEET_NAME", COLOR_INDEX
Dim nameColor As New Dictionary
nameColor.Add "Sheet1", 37
nameColor.Add "Sheet2", 13
nameColor.Add "Sheet3", 6
nameColor.Add "Sheet4", 42

Dim today As Date
today = Date

' Parse through each sheet
For i = 2 To Worksheets.Count
    Worksheets(i).Activate
    
    ' Parse through each row until hitting an empty cell in the A column
    For j = 3 To Worksheets(i).Rows.Count
        ' If the due day is within seven days, of the same month and is not already completed
        If DateDiff("d", Range("C" & CStr(j)).Value, today) <= 7 And DateDiff("m", Range("C" & CStr(j)).Value, today) _
        = 0 And IsEmpty(Range("D" & CStr(j))) = True Then
            Call AddTaskToUpcoming(CStr(j), nameColor)
        End If
        
        If IsEmpty(Range("A" & CStr(j))) = True Then
            Exit For
        End If
    Next j
Next i

End Sub

' Add the given task to the upcoming tasks sheet
Sub AddTaskToUpcoming(rowNumber As String, nameColor As Dictionary)

Dim rowColor As String
rowColor = nameColor.Item(ActiveSheet.Name)

Dim firstEmptyRow As Integer
firstEmptyRow = FindFirstEmptyRow()

Worksheets("Upcoming").Range("A" & firstEmptyRow, "F" & firstEmptyRow).Value = _
ActiveSheet.Range("A" & rowNumber, "F" & rowNumber).Value

Worksheets("Upcoming").Range("A" & firstEmptyRow, "F" & firstEmptyRow).Interior.ColorIndex = rowColor

End Sub

' Finds the first empty row in a sheet
Function FindFirstEmptyRow() As String

For i = 1 To Worksheets("Upcoming").Rows.Count
    If IsEmpty(Worksheets("Upcoming").Range("A" & CStr(i))) = True Then
        FindFirstEmptyRow = i
        Exit For
    End If
Next i

End Function

' Sort all the collected tasks in the upcoming tasks sheet
Sub SortByDueDate()

Dim firstEmptyRow As String
firstEmptyRow = CStr(CInt(FindFirstEmptyRow - 1))

Worksheets("Upcoming").Sort.SortFields.Clear

Range("A1", "F" & firstEmptyRow).Sort Key1:=Range("C1"), Header:=xlYes

End Sub
