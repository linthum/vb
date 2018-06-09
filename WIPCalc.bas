Sub Macro1()

Dim i As Integer
Dim BudgetDate As Date
Dim counter As Integer
Dim dateDiff As Integer
Dim ColDateBudget, ColDateEnd As Integer

'''''##############Budgeting Column############'''''''''
i = 5 ' Column E
ColDateBudget = i
'''''''''''#################################'''''''''''''''''


ActiveSheet.Range("D5") = "Total Hours"
ActiveSheet.Range("D6").FormulaR1C1 = "=sum(r6c4:r25c4)"


BudgetDate = Sheets("test").Range("E6")

dateDiff = Application.WorksheetFunction.Days("06/10/2018", "05/29/2018")

For counter = 1 To dateDiff + 1
    ActiveSheet.Cells(6, i) = BudgetDate + counter
    ActiveSheet.Cells(5, i) = Format(Cells(5, i), "dddd")
    
    'Debug.Print BudgetDate + 1
    i = i + 1
    ColDateEnd = i
Next counter

For i = 7 To 25
    ActiveSheet.Cells(i, 4).FormulaR1C1 = "=sum(r" & i & "c5:r" & i & "c" & ColDateEnd - 1 & ")"
Next i


End Sub

'CALLS THE WIP BASED ON BILL1,2 AND 3
Sub wip(ByVal wip1 As Double, ByVal wip2 As Double, ByVal wip3 As Double)
    
    Sheets("Budget").Unprotect
    'Sheets("Budget").Activate
'    Call weeklySum
   
    Dim rowBudget, rowEnd, i As Integer
    Dim sum As Double
'    Dim wipCollection As Collection
    Dim item As Variant

'    wipCollection = New Collection
'
'    wipCollection.Add wip1
'    wipCollection.Add wip2
'    wipCollection.Add wip3

'    If (Range("c16").Value = "" Or Range("C17").Value = "" Or Range("C21").Value = "") Then
'    Else
'
'Sheets("weekly").PivotTables("weeklyPivot").RefreshTable
    
        rowBudget = 3
        rowEnd = Sheets("Weekly").Range("I" & Rows.Count).End(xlUp).Row
''        rowEnd = dateDiff("d", DateValue(Range("C16")), DateValue(Range("C17")), vbMonday, vbUseSystem)
'
    
        If rowEnd = 1048576 Then

        Else
            Worksheets("Weekly").PivotTables("weeklyPivot").RefreshTable
            Sheets("Summary").Unprotect
            Worksheets("Summary").PivotTables("AuditPivotTable").RefreshTable
            Sheets("Summary").Protect
            sum = 0
            For i = 0 To rowEnd
                sum = Sheets("Weekly").Cells(rowBudget + i, "i").Value + sum
                If wip1 = 0 Then
                    Sheets("Budget").Range("D23").Value = "N/A"
                    Exit For
                ElseIf wip1 < sum Then
                    Sheets("Budget").Range("D23").Value = Sheets("Weekly").Cells(rowBudget + i, "h").Value
                    Exit For
                End If
    
            Next i
    '        End With
            sum = 0
            For i = 0 To rowEnd
                sum = Sheets("Weekly").Cells(rowBudget + i, "i").Value + sum
                If wip2 = 0 Then
                    Sheets("Budget").Range("D24").Value = "N/A"
                    Exit For
                ElseIf wip2 < sum Then
                    Sheets("Budget").Range("D24").Value = Sheets("Weekly").Cells(rowBudget + i, "h").Value
                    Exit For
                End If
    
            Next i
    
            sum = 0
            For i = 0 To rowEnd
                sum = Sheets("Weekly").Cells(rowBudget + i, "i").Value + sum
                If wip3 = 0 Then
                    Sheets("Budget").Range("D25").Value = "N/A"
                    Exit For
                ElseIf wip3 < sum Then
                    Sheets("Budget").Range("D25").Value = Sheets("Weekly").Cells(rowBudget + i, "h").Value
                    Exit For
                End If
    
            Next i

        End If
 

'Debug.Print "wip1.......", wip1
'Debug.Print "wip1.......", wip2
'Debug.Print "wip1.......", wip3

Sheets("Budget").Protect

End Sub