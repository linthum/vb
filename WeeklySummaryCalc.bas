Sub WeeklySummaryCalculate()
'+-----------------------------------+
'|   Calculates the weekly summary   |
'+-----------------------------------+
    Dim SelectedDate, DateToCheck, DateToUpdate, StartDate, EndDate As Date
    Dim DaysBetween, StartCol, EndCol, PrintCol, WklyStartCol, WklyEndCol, WorkingCol, ValueToUpdate, ValueToAdd, CountOfStaff, WklyStartRow, WklyEndRow As Integer
    Dim StaffName As String
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim WeeklySummaryTab, CurrentWorkSheet As Worksheet
        
    
    
    Set WeeklySummaryTab = wb.Sheets("Weekly Summary")
   
    
    Application.ScreenUpdating = False
    WeeklySummaryTab.Select
    
    If WeeklySummaryTab.Range("D2").Text = "" Then
        WklyStartCol = 3 '2
        WklyEndCol = 2
    Else
        Range("C2").Select
        Selection.End(xlToRight).Select
        WklyStartCol = 3 ' 2
        WklyEndCol = Selection.Column
    End If
    
    '+----------------------------------------------+
    '|   Gets the number of staff to loop through   |
    '+----------------------------------------------+
    If WeeklySummaryTab.Range("A4").Text = "" Then
        CountOfStaff = 1
    Else
        Range("A3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlDown)).Select
        CountOfStaff = Selection.Count
    End If
    
    
    WklyStartRow = 3
    WklyEndRow = WklyStartRow + CountOfStaff
    
    
    '+-----------------------------+
    '|   Loops through the staff   |
    '+-----------------------------+
    For b = WklyStartRow To WklyEndRow - 1
        
        StaffName = Left(Replace(Cells(b, 1).Text, "'", ""), 30)
        
        
        Set CurrentWorkSheet = wb.Sheets(StaffName) ' To Change
        CurrentWorkSheet.Select
    
        Range("E6").Select 'Start Date
        Selection.End(xlToRight).Select
        EndCol = Selection.Column
     
        '+-------------------------------------------------+
        '|   Clears contents so old values are not added   |
        '+-------------------------------------------------+
        WeeklySummaryTab.Select
        Cells(b, WklyStartCol).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.ClearContents
        
        '+-------------------------------------------------+
        '|   Loops through each column in the staff's tab  |
        '|   and extracts the hours for that day and adds  |
        '|   it to the summary                             |
        '+-------------------------------------------------+
        For i = EndCol To 5 Step -1 '5 Step -1
            SelectedDate = CurrentWorkSheet.Cells(6, i).Value
            
            For a = WklyEndCol To WklyStartCol Step -1
                DateToCheck = WeeklySummaryTab.Cells(2, a).Value
                DaysBetween = SelectedDate - DateToCheck
                If (DaysBetween) <= 6 Then
                    
                    WorkingCol = a
                End If
            Next a
            'MsgBox SelectedDate & " Falls under  " & DateToUpdate
            ValueToUpdate = WeeklySummaryTab.Cells(b, WorkingCol).Value
            ValueToAdd = CurrentWorkSheet.Cells(26, i).Value
            ValueToUpdate = ValueToUpdate + ValueToAdd
            WeeklySummaryTab.Cells(b, WorkingCol).Value = ValueToUpdate
        Next i

    Next b
    
    '+-----------------------------+
    '|   Adds timestamp to Cell A1 |
    '+-----------------------------+
    WeeklySummaryTab.Range("A1").Value = "Last updated: " & vbNewLine & Now()
    
    
    Columns("A:B").EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub

Sub AddStaffToWeekly(ByVal StaffName As String)
'+--------------------------------------+
'|   Adds staff to the weekly summary.  |
'|   Called at from the CreateNewSheet  |
'|   function in the fees module        |
'+--------------------------------------+
    Dim WorkingRow As Integer
    If Worksheets("Weekly Summary").Range("A3").Value = "" Then
        Worksheets("Weekly Summary").Range("A3").Value = StaffName
        WorkingRow = 3
    ElseIf Worksheets("Weekly Summary").Range("A4").Value = "" Then
        Worksheets("Weekly Summary").Range("A4").Value = StaffName
        WorkingRow = 4
    Else
        Sheets("Weekly Summary").Select
        Range("A3").Select
        Selection.End(xlDown).Select
        WorkingRow = Selection.Row
        WorkingRow = WorkingRow + 1
        Worksheets("Weekly Summary").Cells(WorkingRow, 1).Value = StaffName
        
    End If
        Worksheets("Weekly Summary").Cells(WorkingRow, 2).Formula = "=VLOOKUP(" & Cells(WorkingRow, 1).Address & ",Staff_Fees!$C:$F,2,FALSE)"
    Sheets("Weekly Summary").Select
    Columns("A:B").EntireColumn.AutoFit
    
    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

End Sub




Sub TriggerFunc()
'+---------------------+
'|   Testing Function  |
'+---------------------+
    Call WeeklySummarySetup("18/06/2018", "14/08/2018")
    'Call AddStaffToWeekly("Eimear McCarthy")
    'Call WeeklySummaryCalc
End Sub

Sub WeeklySummarySetup(ByVal StartDate As Date, ByVal EndDate As Date)
'+-------------------------------------+
'|   Creates the weekly summary sheet  |
'+-------------------------------------+
    Dim SheetName, DayCheck, DayName As String
    Dim Days, CurrentCol, StartCol, EndCol As Integer
    CurrentCol = 3 ' 2 '#############################################################################
    
    'Days = CDate(EndDate) - CDate(StartDate)
    SheetName = "Weekly Summary" ' & Sheets.Count() 'Testing only
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SheetName
    Worksheets(SheetName).Select
    
    '+-------------------------------------------------------+
    '|   Adds all the dates within the date range selected   |
    '+-------------------------------------------------------+
    Cells(1, CurrentCol).Select
    ActiveCell.FormulaR1C1 = "=TEXT(R[1]C,""DDDDD"")"
    Cells(2, CurrentCol).Select
    
    'Cells(2, 2).Value = StartDate
    'MsgBox Weekday(StartDate, 2)
'    Cells(2, CurrentCol).Value = StartDate - Weekday(StartDate, 2) + 6
'    Days = CDate(EndDate) - CDate(StartDate - Weekday(StartDate, 2) + 6)
    
    DayName = Format(StartDate, "dddd")
    'MsgBox DayName
    If DayName = "Sunday" Then
        Cells(2, CurrentCol).Value = (StartDate - 7) - Weekday(StartDate, vbUseSystem) + 2
        Days = CDate(EndDate) - CDate((StartDate - 7) - Weekday(StartDate, vbUseSystem) + 2)
    Else
        Cells(2, CurrentCol).Value = StartDate - Weekday(StartDate, vbUseSystem) + 2
        Days = CDate(EndDate) - CDate(StartDate - Weekday(StartDate, vbUseSystem) + 2)
    End If
    
    
    
    
    For i = 1 To Days
        Cells(2, CurrentCol + i).Select
        ActiveCell.FormulaR1C1 = "=RC[-1]+1"
        Cells(1, CurrentCol + i).Select
        ActiveCell.FormulaR1C1 = "=TEXT(R[1]C,""DDDDD"")"
    Next i
    
    '+-----------------------------------------------------------------------+
    '|   Selects Row 1 and Copies it. Then pastes the values back in Row 1   |
    '+-----------------------------------------------------------------------+
    Rows("1:2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    '+-------------------------------------------------------------------------------------+
    '|   Deletes all columns which are not a monday. Also does not delete the first date   |
    '+-------------------------------------------------------------------------------------+
    StartCol = 4
    Cells(1, CurrentCol).Select
    Selection.End(xlToRight).Select
    EndCol = Selection.Column
    For i = EndCol To StartCol Step -1
       ' Cells(1, i).Value.Select
        If Cells(1, i).Value <> "Monday" Then
             Columns(i).Select
             Selection.Delete Shift:=xlToLeft
        End If
    Next i
    
    
    
    Cells.Select
    Selection.Columns.AutoFit
End Sub

