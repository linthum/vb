'formats the Budget sheet with number and currency formatting
Public formatSheetCounter As Integer
Public Property Get getFormatSheetCounter() As Integer
    getFormatSheetCounter = formatSheetCounter
End Property


Sub formatBudgetSheet(lastRow As Integer)
formatSheetCounter = 1

  Sheets("Budget").Activate
    
    ActiveWindow.Zoom = 90

    Range("E5:I5").Select
    Selection.Style = "Total"
    Selection.Style = "20% - Accent1"
    Selection.Font.Bold = True
    
    Range("E6:I" & lastRow).Select
    Selection.Style = "Note"
    Selection.Font.Bold = True
    
    Range("I6:I" & lastRow).Select
    Selection.Style = "Currency"
    
    Range("G6:G" & lastRow).Select
    Selection.Style = "Currency"
    
    Range("C23:C25").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_-$* #,##0.0_-;-$* #,##0.0_-;_-$* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""??_-;_-@_-"
    Selection.Font.Bold = True
        
    Range("C21").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_-$* #,##0.0_-;-$* #,##0.0_-;_-$* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""??_-;_-@_-"
    Selection.Font.Bold = True
    
    Range("C19").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_-$* #,##0.0_-;-$* #,##0.0_-;_-$* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""??_-;_-@_-"
    Selection.Font.Bold = True
    
    Range("C16:C17").Select
    Selection.Font.Bold = True
    
    formatSheetCounter = 2
    
End Sub

 