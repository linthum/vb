Sub testMacro()

'Call CreateNewSheet("test_xyz", "01/05/2018", "08/05/2018")
'Call formatNewSheet(14)

End Sub
'LOCKS THE NEW SHEET
Sub lockingCells()

ActiveSheet.Protect

End Sub

'FORMATES THE NEW SHEET WHENEVER IT IS CREATED
Sub formatNewSheet(ByVal lastCol As Integer)

Dim borderRange As Range

Set borderRange = Range(Cells(5, 4), Cells(25, lastCol))

    
    ActiveWindow.Zoom = 90
    Cells.EntireColumn.AutoFit
    ActiveWindow.DisplayGridlines = False
    borderRange.Font.Bold = True
    borderRange.HorizontalAlignment = xlCenter
    
    Range("B2:C3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    With Range("B2:C3")
        .Font.Size = 18
        .Cells.Style = "Heading 1"
    End With
    
    
    
    Range("B5:B6").Select
    Range("B6").Activate
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C5:C6").Select
    Range("C6").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Range("D5").Select
    Selection.Font.Bold = True
    
    '' Give your border range
    borderRange.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    

Range(Cells(26, "d"), Cells(26, lastCol)).Style = "Total"

With Range(Cells(7, 5), Cells(25, lastCol)).Validation
    .Delete
    .Add Type:=xlValidateDecimal, _
    AlertStyle:=xlValidAlertStop, _
    Formula1:="0", Formula2:="24"
    .InputTitle = "Integers"
    .ErrorTitle = "Integers"
    .InputMessage = "Enter an integer from 0 to 24 only!"
    .ErrorMessage = "You must enter a number"
End With

Range(Cells(7, 5), Cells(25, lastCol)).Locked = False
Range(Cells(7, 5), Cells(25, lastCol)).Value = 0
    
Range("E5").Select
ActiveWindow.FreezePanes = True
ActiveSheet.ScrollArea = Range(Cells(1, 1), Cells(28, lastCol)).Address
Call lockingCells

'Call weeklySum(getLastCol(lastCol))
    
End Sub