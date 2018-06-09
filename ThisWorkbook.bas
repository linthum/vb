Public lastSheetName As String
Public sheetDeleted As Boolean


'Hides the ribbon when workbook is open
Private Sub Workbook_Open()
'Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
' CommandBars.ExecuteMso "MinimizeRibbon"
If Application.CommandBars("Ribbon").Height >= 150 Then
    SendKeys "^{F1}"
End If


End Sub

'Private Sub Workbook_NewSheet(ByVal Sh As Object)
'Application.DisplayAlerts = False
'If (StrComp(Sh.Name, "sheet", vbTextCompare)) = 1 Then
'    MsgBox (True)
'    MsgBox (Sh.Name)
'End If

'ActiveSheet.Delete


'End Sub

'Private Sub Workbook_Open()
'
''alertTime = Now + TimeValue("00:00:10")
''Application.OnTime alertTime, "summarySheet"
'
'
'End Sub


Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)

    Dim lastRow, i As Integer

    If (Sh.Name = "Summary") Then
    
    Else
'        Call weeklySum
    
    
    End If
    
    If fee.btnFeeBreakCounter = 2 Then
        
    Else
        lastRow = Sheets("Budget").Range("E" & Rows.Count).End(xlUp).Row
        'MsgBox (Sh.Name)
        For i = 6 To lastRow
        '    Debug.Print Sheets("Budget").Cells(i, "E").Value
            If Sheets("Budget").Cells(i, "E").Value = Sh.Name Then
                    
    '            Debug.Print Range(Cells(i, "e"), Cells(i, "i")).Address
                Sheets("Budget").Unprotect
                Sheets("Budget").Activate
                
'                If (i = 6) Then
                Range(Cells(i, "e"), Cells(i, "i")).Select
                Sheets("Budget").Unprotect
                Selection.Delete Shift:=xlUp
'                Else
'                  Sheets("Budget").Range(Cells(i, "e"), Cells(i, "i")).Delete Shift:=xlShiftUp
'                End If
                
                
              
                
                Call weeklySum
                Call summarySheet
                    
            End If
            
        Next i
        
        btnFeeBreakCounter = 4
        sheetDeleted = True
    
        Sheets("Budget").Activate
        
    End If
    
    
'
'Call weeklySum
'Call summarySheet
'Call feeBreakDown

End Sub


Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
'MsgBox ("i'm treiggered....." & Sh.Name)
'lastSheetName = Sh.Name
End Sub
Public Property Get getLastSheetName() As String
    getLastSheetName = lastSheetName
End Property

Public Property Get getSheetDeleted() As Boolean
    getSheetDeleted = sheetDeleted
End Property