
'Add new staff button on the Budget sheet
'needs to be disabled initially
'
Sub btnAdd_CreateSheet()


Dim SheetName As String
Dim BudgetDate As String
Dim EndDate As String
Dim lastRow As Integer
Dim wks As Variant
Dim sheetFound As Boolean

sheetFound = False
SheetName = ActiveSheet.Range("F3").Value
BudgetDate = ActiveSheet.Range("C16").Value
EndDate = ActiveSheet.Range("C17").Value

For Each wks In Worksheets
    If wks.Name = Sheets("Budget").Range("F3").Value Then
        sheetFound = True
        Exit For
    End If
Next wks

If sheetFound = False Then
    Call CreateNewSheet(SheetName, BudgetDate, EndDate)
    Call weeklySum
    Call summarySheet
    Call feeBreakDown
Else
    MsgBox ("sheet there dude")
End If

End Sub