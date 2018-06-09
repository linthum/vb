'Public deletedSheet As Boolean

Private Sub BtnAddStaff_Click()
'FrmSelectStaff.Unload
FrmSelectStaff.Show

End Sub

'
'Private Sub BtnAddStaff_Click()
'    Unload FrmSelectStaff
'    FrmSelectStaff.Show
'End Sub

Private Sub Worksheet_Activate()

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True '##################

'Dim wks As Variant
'Dim sheetFound As Boolean
'
'sheetFound = False
'
'For Each wks In Worksheets
'    If wks.Name = "Summary" Then
'        sheetFound = True
'        Exit For
'    End If
'Next wks

'If sheetFound = False Then
'Else
'    Application.ScreenUpdating = False
'
'    If fee.getFeeBreakCounter = 2 Or fee.getFeeBreakCounter = 1 Then
'
'    Else
'        Sheets("weekly").PivotTables("weeklyPivot").RefreshTable
'
'        Sheets("Summary").Unprotect
'        Sheets("Summary").PivotTables("AuditPivotTable").RefreshTable
'        Sheets("Summary").Protect
'    End If
'
'End If

'Dim wks As Variant


'deletedSheet = False
'
'For Each wks In Worksheets
'    If wks.Name = ThisWorkbook.getLastSheetName Then
'        deletedSheet = True
'        Exit For
'    End If
'

'Next wks

'    Debug.Print "i'm in Budget sheeet......", fee.getFeeBreakCounter
            
'        Debug.Print "wip triggered"

End Sub

'Calculates the all the bills and outputs the results on the Budget sheet

Private Sub Worksheet_Calculate()


'    Debug.Print "deleted sheet...", deletedSheet
'
'    Application.ScreenUpdating = False
'
'    If fee.getFeeBreakCounter = 4 Or ThisWorkbook.getSheetDeleted = True Then
'        Call feeBreakDown
'        Call weeklySum
'        Call summarySheet
'
'
'    End If
  
    
    
    
    
    'Debug.Print "horaaaa"
'    If fee.getFeeBreakCounter <> 0 Then
         
'            Else
'
'            Debug.Print wip1
'            Debug.Print wip2
'            Debug.Print wip3
'
'            End If
            
            
' Sheets("Budget").Protect
'
End Sub


Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    'This function is called when someone clicks on one of the hyperlinks on the "Budget Sheet"
    'It works by taking taking the value of whatever cell the hyperlink is referencing and using it to search for the correct tab
    'It then creats a copy of that tab and saves it in a new notebook
    'The value is then also used in the files name along with the date when saving
    
    
    'MsgBox CStr(ActiveCell.Value)
    Dim TabToExtract As String

    TabToExtract = ActiveCell.Value
    TabToExtract = Left(TabToExtract, 30)
    TabToExtract = Replace(TabToExtract, "'", " ")


    Dim FName           As String
    Dim FPath           As String
    Dim NewBook         As Workbook

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing


   
    FName = TabToExtract & " (" & Format(Date, "dd-mm-yy") & ").xls"

    Set NewBook = Workbooks.Add

    ThisWorkbook.Sheets(TabToExtract).Copy Before:=NewBook.Sheets(1)

    Application.DisplayAlerts = False
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet3").Select
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True


    If Dir(GetFolder & "\" & FName) <> "" Then
        MsgBox "File " & flGetFolderdr & "\" & FName & " already exists"
    Else
        NewBook.SaveAs FileName:=GetFolder & "\" & FName
    End If


End Sub

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'With ActiveWindow


'    'This function is used to updated the billing figures
        
                        
'                        If PercentFee >= 0.15 Then
'                            Range("C23").Value = Fees * 0.5
'                            Range("A23").Value = "Interm 1 Bill at 50%"
                            
'                        Else
'                            Range("C23").Value = "N/A"
'                        End If
                 
                    
'                        If PercentFee >= 0.15 Then
'                            Range("C23").Value = Fees * 0.35
'                            Range("A23").Value = "Interm 1 Bill at 35%"
''                        Else
'                            Range("C23").Value = "N/A"
'                        End If
'                    Else
'                        Range("C23").Value = "N/A"
'                    End If
'
'                    If Fees > 50000 Then
'                        If PercentFee >= 0.5 Then
'                            Range("C24").Value = Fees * 0.35
'                            Range("A24").Value = "Interm 2 Bill at 35%"
'                        Else
'                            Range("C24").Value = "N/A"
'                        End If
'                    Else
'                        Range("C24").Value = "N/A"
'                    End If

'
'                    If Fees < 15000 Then
'                        If PercentFee >= 0.4 Then
'                            Range("C25").Value = Fees
'                            Range("A25").Value = "Final Bill at 100%"
'                        Else
'                            Range("C25").Value = "N/A"
'                        End If
'                    ElseIf Fees > 15000 And Fees < 50000 Then
'                        If PercentFee >= 0.75 Then
'                            Range("C25").Value = Fees * 0.5
'                            Range("A25").Value = "Final Bill at 50%"
'                        Else
'                            Range("C25").Value = "N/A"
'                        End If
'                    ElseIf Fees > 50000 Then
'                        If PercentFee >= 0.75 Then
'                            Range("C25").Value = Fees * 0.3
'                            Range("A25").Value = "Final Bill at 30%"
'                        Else
'                            Range("C25").Value = "N/A"
'                        End If
'                    Else
'                        Range("C25").Value = "N/A"
'                    End If
'             End If
'        End If
        
        
        '################### WIP CALCULATION ##########''''''''''
'
'        Dim wipAmt As Long
'
'
'
'
'        End With
'
'
'
'
'End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)



Application.ScreenUpdating = False

If MoreFormatting.formatSheetCounter <> 1 And IsEmpty(Range("C19")) = False Then


            Dim Fees As String
            Dim PercentFee As Double
            Dim ErrorCheck As Boolean
            Dim wip1, wip2, wip3 As Double
            Dim bill_1, bill_2, bill_3 As Double
            Dim cost As Double
            Dim StrInterm_1, StrInterm_2, StrFinal As String

            bill_1 = 0
            bill_2 = 0
            bill_3 = 0
            cost = 0

            StrInterm_1 = "Interm 1 Bill"
            StrInterm_2 = "Interm 2 Bill"
            StrFinal = "Final Bill"


            ErrorCheck = False
            Fees = Worksheets("Budget").Cells(19, 3)

    '        if Sheets("Budget").Range("C21").Value

            '' problem with clearing sheet
'            if (Sheets("Budget").range("C21").value


            cost = Sheets("Budget").Range("C21").Value

            If Fees <> "" Then
                If IsError(Cells(22, 3)) Then
                    ErrorCheck = True
                End If

                If Range("C22").Text <> "" And ErrorCheck <> True Then
                    PercentFee = Range("C22").Value

                    If Fees >= 15000 And Fees <= 50000 Then
                        bill_1 = Fees * 0.5
                        bill_2 = 0
                        bill_3 = Fees * 0.5

                        wip1 = cost * 0.15
                        wip2 = 0
                        wip3 = cost * 0.75

                        StrInterm_1 = "Interim 1 Bill:   50% of Expected Fees"
                        StrInterm_2 = "Interim 2 Bill:   N/A"
                        StrFinal = "Final Bill:   50% of Expected Fees"


                    ElseIf Fees > 50000 Then
                        bill_1 = Fees * 0.35
                        bill_2 = Fees * 0.35
                        bill_3 = Fees * 0.3

                        wip1 = cost * 0.15
                        wip2 = cost * 0.5
                        wip3 = cost * 0.75

                        StrInterm_1 = "Interim 1 Bill:   35% of Expected Fees"
                        StrInterm_2 = "Interim 2 Bill:   35% of Expected Fees"
                        StrFinal = "Final Bill:   30% of Expected Fees"

                    ElseIf Fees < 15000 Then
                        bill_1 = 0
                        bill_2 = 0
                        bill_3 = Fees

                        wip1 = 0
                        wip2 = 0
                        wip3 = cost * 0.4

                        StrInterm_1 = "Interim 1 Bill:   N/A"
                        StrInterm_2 = "Interim 2 Bill:   N/A"
                        StrFinal = "Final Bill:   100% of Expected Fees"
                    Else

                    End If


                End If
            End If

            Sheets("Budget").Unprotect

            With Sheets("Budget")
                    .Range("C23").Value = bill_1
                    .Range("C24").Value = bill_2
                    .Range("C25").Value = bill_3
                    .Range("A23").Value = StrInterm_1
                    .Range("A24").Value = StrInterm_2
                    .Range("A25").Value = StrFinal
            End With

'            Sheets("Budget").Protect


         Call wip(wip1, wip2, wip3)
    Else
    End If

End Sub