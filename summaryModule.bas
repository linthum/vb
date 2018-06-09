Option Explicit

'creates a sheet where task and number of hours of each person are recorded. and calls the summary pivot table
Sub summarySheet()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim wks As Worksheet, strName As String
Dim workSheetArr As Collection
Dim i As Variant
Dim BudgetCol, EndCol As Integer
Dim vlk_StaffName As String
Dim vlk_BudgetRange As Integer
Dim vlk_EndRange As Integer
Dim rowEnd As Integer
Dim vlk_Range As Range
Dim BudgetLoop As Integer
Dim collCounter As Integer
Dim j As Integer
Dim vlk_Result As Variant
Dim pivotRange As Range
Dim clearRange As Range
Dim taskRange As Range

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False
Sheets("DSheet").UsedRange.Clear
'Application.ScreenUpdating = False

Set workSheetArr = New Collection
Sheets("DSheet").Range("A1") = "Task"
Sheets("DSheet").Range("B1") = "Sub-Task"
Sheets("DSheet").Range("C1") = "Staff Name"
Sheets("DSheet").Range("D1") = "Hours"
'Sheets("DSheet").Range("F1") = "Task"

Set taskRange = Worksheets("Data").Range("A1:B19")
         
Dim cell As Variant
Dim sumRange As Range

Set sumRange = Range("D7:C25")

For Each wks In Worksheets
        If VBA.StrComp(wks.Name, "Budget", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Staff_Fees", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Client_Codes", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "DSheet", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Data", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Weekly", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Instructions", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Summary", vbTextCompare) = 0 Or _
        VBA.StrComp(wks.Name, "Group Fee Billing Schedule", vbTextCompare) = 0 Or _
         VBA.StrComp(wks.Name, "Weekly Summary", vbTextCompare) = 0 Then
     Else
        workSheetArr.Add wks.Name
     End If
Next wks

Dim counter As Integer

For i = 1 To workSheetArr.Count '######################################################################
    Dim StaffName, TabName As String
    TabName = workSheetArr(i)
    StaffName = Worksheets(TabName).Range("B2").Text
    counter = 2
    Sheets("DSheet").Activate
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    taskRange.Copy
    rowEnd = Worksheets("DSheet").Cells(Rows.Count, "A").End(xlUp).Row
    
    Cells(rowEnd + 1, 1).Select
    
    Sheets("DSheet").Paste
    ' <= 20 because there are 20 subtasks
    Do While (counter <= 20)
    
        With Worksheets("DSheet")
            rowEnd = .Cells(.Rows.Count, "C").End(xlUp).Row
            .Cells(rowEnd + 1, 3).Value = StaffName  ' workSheetArr(i) '##########################################
            .Cells(rowEnd + 1, 4).Formula = "=VLOOKUP(B" & rowEnd + 1 & ",'" & workSheetArr(i) & _
                                            "'!" & sumRange.Address & ",2,FALSE)"
        End With
        
        counter = counter + 1
        rowEnd = rowEnd + 1
    Loop

Next i

rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
Set pivotRange = Worksheets("DSheet").Range(Cells(1, 1), Cells(rowEnd, 4))
'
Debug.Print pivotRange.Address
'
Call createPivot(workSheetArr, pivotRange)
End Sub


'creates a pivot table in the same weekly sheet

Sub createPivot(SheetNames As Collection, pivotRange As Range)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim PivotSheet, DataSheet As Worksheet

Sheets("Summary").Visible = True
Sheets("Summary").Select
Sheets("Summary").Unprotect
ActiveSheet.UsedRange.Clear

Set PivotSheet = Worksheets("Summary")
Set DataSheet = Worksheets("DSheet")

PivotSheet.Select
    
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=pivotRange, Version:=6). _
            CreatePivotTable TableDestination:=PivotSheet.Cells(4, 2), _
            TableName:="AuditPivotTable", DefaultVersion:=6
        
    PivotSheet.Activate
    
    
    With ActiveSheet.PivotTables("AuditPivotTable").PivotFields("Task")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    'Party begins here
    With ActiveSheet.PivotTables("AuditPivotTable").PivotFields("Sub-Task")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables("AuditPivotTable").PivotFields("Staff Name")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("AuditPivotTable").AddDataField ActiveSheet.PivotTables _
        ("AuditPivotTable").PivotFields("Hours"), "Sum - Hours", xlSum
    
    
    'Ends here
         
    With ActiveSheet.PivotTables("AuditPivotTable")
        .ColumnGrand = True
        .RowGrand = True
    End With
    
    ActiveSheet.PivotTables("AuditPivotTable").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("AuditPivotTable").TableStyle2 = "PivotStyleMedium7"
    
    ActiveWindow.DisplayGridlines = False
    Range("D5").Select
    ActiveWindow.FreezePanes = True
    Rows("1:3").EntireRow.Hidden = True
    ActiveSheet.Protect
    
End Sub

