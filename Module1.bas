Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("C27").Select
    ActiveCell.FormulaR1C1 = "v"
    Range("C28").Select
End Sub