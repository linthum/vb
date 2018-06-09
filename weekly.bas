

'THIS FUNCTION CREATES A SHEET WEEKLY WHICH WILL RECORD THE NAME, HOURS, COST AND DATE OF THE PERSON.
'VERY IMPORTANT FUNCTION TO TRIGGER THE WIP
Sub weeklySum()

Application.ScreenUpdating = False

Dim wks As Variant
Dim staffNames_Arr As Collection
Dim lastCol As Integer
Dim rowEnd As Integer
Dim sumRange, criteriaRange1, criteriaRange2 As Range
Dim BudgetDate As String
Dim EndDate As String
Dim nDays As Integer
Dim StaffName As String
Dim counter As Integer
Dim i As Variant
Dim lastRow As Integer
Dim weeklyPivotRange As Range

Set staffNames_Arr = New Collection

Sheets("Weekly").Activate
Sheets("Weekly").Visible = True
Sheets("Weekly").Select
ActiveSheet.Unprotect
ActiveSheet.UsedRange.Clear

ActiveSheet.Range("A1").Value = "Staff Name"
ActiveSheet.Range("B1").Value = "Daily Date"
ActiveSheet.Range("C1").Value = "Daily Hours"
ActiveSheet.Range("D1").Value = "Daily Cost"


BudgetDate = CDate(Sheets("Budget").Range("C16").Value)
EndDate = CDate(Sheets("Budget").Range("C17").Value)

nDays = dateDiff("d", BudgetDate, EndDate, vbUseSystemDayOfWeek, vbUseSystem) + 1
lastCol = nDays + 1 + 3 ' empty column + B C D

For Each wks In Worksheets
    counter = 0
    i = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        If VBA.StrComp(wks.Name, "Budget", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Staff_Fees", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Client_Codes", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "DSheet", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Data", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Weekly", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Instructions", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Group Fee Billing Schedule", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Summary", vbTextCompare) = 0 Then
     Else
'        staffName = WorksheetFunction.Trim(wks.Name) '=========================================================
         StaffName = wks.Range("B2").Text
         Dim SheetName As String
         SheetName = WorksheetFunction.Trim(wks.Name)
        'ActiveSheet.Range("H3").Value = staffName
        staffNames_Arr.Add StaffName
        
        Do While (counter < nDays)
            ActiveSheet.Cells(i, 1).Value = StaffName
'            ActiveSheet.Cells(i, 2).Value = Format(BudgetDate) + counter, "dd/mm/yyyy")
             ActiveSheet.Cells(i, 2).Value = CDate(BudgetDate) + counter
            
            ActiveSheet.Cells(i, 3).Formula = "=IFERROR(SUM('" & wks.Name & "'!" & _
            Range(Cells(7, 5 + counter), Cells(25, 5 + counter)).Address & "),0)" '=============================
             
            Cells(i, "d").Formula = "=IFERROR((VLOOKUP(" & Cells(i, "a").Address & ",Staff_Fees!$C$1:$F$744,4,FALSE)*Weekly!" & Cells(i, "c").Address & "),0)"
            counter = counter + 1
            i = i + 1
            lastRow = i - 1
        Loop
     End If
Next wks

Sheets("weekly").Visible = True
   Set weeklyPivotRange = Range(Cells(1, 1), Cells(lastRow, "d"))
    Call createWeeklyPivot(weeklyPivotRange)
End Sub


'CREATES A PIVOT OF THE WEEKLY TABLE IN THE SAME SHEET
Sub createWeeklyPivot(pivotRange As Range)

Sheets("Weekly").Activate
Dim pivotName As String

pivotName = "weeklyPivot"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange, _
    Version:=6).CreatePivotTable TableDestination:= _
        ActiveSheet.Cells(2, "h"), TableName:=pivotName, DefaultVersion:=6
    
    With ActiveSheet.PivotTables(pivotName).PivotFields("Daily Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables(pivotName).AddDataField ActiveSheet.PivotTables( _
        pivotName).PivotFields("Daily Cost"), "Sum - Daily Cost", xlSum
    
    ActiveSheet.PivotTables(pivotName).RowAxisLayout xlTabularRow
    
    With ActiveSheet.PivotTables(pivotName)
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    ActiveSheet.PivotTables(pivotName).PivotFields("Staff Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivotName).PivotFields("Daily Date").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivotName).PivotFields("Daily Hours").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivotName).PivotFields("Daily Cost").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)

    Sheets("weekly").Visible = False
End Sub