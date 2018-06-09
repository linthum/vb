Option Explicit
Public btnFeeBreakCounter As Integer
Sub Initial_Details_Button()
    Unload StartBudget
    
    StartBudget.Show
    
End Sub


Sub BtnAddStaff_Click()
    Unload FrmSelectStaff
    FrmSelectStaff.Show
End Sub

'clear the tool when Reset button is pressed
'detelets the user created sheet expcept the Summary, Weekly and Budget sheet.
'DO NOT DELTE SUMMARY AND WEEKLY SHEET!!' 'ONLY HIDE/UNHIDE THEM'
'ALSO CLEARS THE Budget SHEET, DO NOT CLEAR THE COLUMN, ONLY USE CLEARCONTENT, LEAVE THE FORMATTING AS IT IS
Sub ClearValues()
    
    btnFeeBreakCounter = 2

    Dim answer As Integer
    Dim i As Integer
    answer = MsgBox("Are you sure you want to reset the tool? Once Budgeted, it cannot be undone." & vbNewLine & vbNewLine & "Reports need to be manualy removed.", vbYesNo + vbQuestion, "Empty Sheet")
    
    If answer = vbYes Then
        'Application.EnableEvents = False
        For i = Worksheets.Count To 1 Step -1
            If Worksheets(i).Name = "Budget" Then
            
            ElseIf Worksheets(i).Name = "Staff_Fees" Then
            
            ElseIf Worksheets(i).Name = "Instructions" Then
            
            ElseIf Worksheets(i).Name = "Client_Codes" Then
            
            ElseIf Worksheets(i).Name = "DSheet" Then
            
            ElseIf Worksheets(i).Name = "Data" Then
            
            ElseIf Worksheets(i).Name = "Summary" Then
            ElseIf Worksheets(i).Name = "Weekly" Then
            ElseIf Worksheets(i).Name = "Group Fee Billing Schedule" Then
            
            
            
            
            Else
                Application.DisplayAlerts = False
                Sheets(i).Delete
                
                
                'ActiveWindow.SelectedSheets.Delete
                Application.DisplayAlerts = True
            End If
        Next i
        
    Sheets("Weekly").UsedRange.Clear

    Sheets("DSheet").UsedRange.Clear
    Sheets("Summary").Visible = False
    Sheets("Budget").Unprotect
    
    
    With Sheets("Budget")
    .Range("C:C").ClearContents
    
    .Range("E6:I" & Rows.Count).ClearContents
        .Range("E6:I" & Rows.Count).ClearFormats
        .Range("D23:D25").ClearContents
    End With
    
    
    Sheets("weekly").Visible = False
    Sheets("Budget").Protect
    
    
    Else
    'do nothing
    End If
End Sub



'THIS FUNCTION WILL PERFORM A VLOOKUP ON THE STAFF NAME TO GET ROLE, CHARGE RATE, HOURS AND COST,
'IT WILL ALSO CALCULATE THE RECOVERY RATE AS WELL
'FUNCTION CALLED AFTER EACH SHEET IS CREATED.
Sub feeBreakDown()
    

    Sheets("Budget").Unprotect
    Sheets("Budget").Activate

    Dim LastStaff As Integer
    Dim i As Integer
    
    LastStaff = Sheets("Budget").Cells(Rows.Count, "E").End(xlUp).Row
    
    If LastStaff = 5 Then
    Else
    Application.ScreenUpdating = False
    
    For i = 6 To LastStaff
        Sheets("Budget").Unprotect
        Dim TabName, StaffSearch As String
        TabName = Left(Replace(Cells(i, "E").Text, "'", ""), 30)
        StaffSearch = Cells(i, "E").Text
        
        With Sheets("Budget")
            .Cells(i, "f") = Application.VLookup(StaffSearch, Sheets("Staff_Fees").Range("C:D"), 2, False) 'Grade      '.Formula = "=IFERROR(VLOOKUP(" & Cells(i, "E").Address & ",Staff_Fees!$C:$F,2,FALSE),0)"
            
            Sheets("Budget").Unprotect
            .Cells(i, "g") = Application.VLookup(StaffSearch, Sheets("Staff_Fees").Range("C:F"), 4, False) '.Formula = "=IFERROR(VLOOKUP(" & Cells(i, "E").Address & ",Staff_Fees!$C:$F,4,FALSE),0)"
            Sheets("Budget").Unprotect
'            Debug.Print Cells(i, "e").Value
            '.Cells(i, "H").Formula = "='" & Sheets("Budget").Cells(i, "e").Value & "'!D6"
            .Cells(i, "H").Formula = "='" & TabName & "'!D6"
            Sheets("Budget").Unprotect
            .Cells(i, "i").Formula = "=" & Cells(i, "g").Address & "*" & Cells(i, "H").Address
        End With
        

    Next i
    Sheets("Budget").Unprotect
    With Sheets("Budget")
'        Debug.Print "=IFERROR(sum(I6:I" & LastStaff & ")," & 0 & ")"
        .Range("C21").Formula = "=IFERROR(sum(I:I" & ")," & 0 & ")"
    End With
    Sheets("Budget").Unprotect
    
    With Sheets("Budget").Range("C22")
        .Formula = "=IFERROR(sum(" & Range("C19").Address & "/" & Range("C21").Address & ")," & 0 & ")"
    'Range("C22").Formula = "=(" & Range("c19").Value & "-" & Range("c21").Value & ")/" & Range("C19").Value
        
        Sheets("Budget").Unprotect
        
        .Style = "Percent"
        
        Sheets("Budget").Unprotect
        
        .NumberFormat = "0.0%"
        
        Sheets("Budget").Unprotect
        .Font.Bold = True
    
    End With
    
    
     btnFeeBreakCounter = btnFeeBreakCounter + 1
    
    Application.ScreenUpdating = False
    
     formatBudgetSheet (LastStaff)
'    Call summarySheet
'    Call weeklySum
    
    Sheets("Budget").Activate
    Sheets("Budget").Range("E5").Select
    Sheets("Budget").Protect
    End If
    
    
   
End Sub

'CREATES A NEW SHEET BASED ON SHEETNAME, BudgetDATE AND ENDDATE.

Sub CreateNewSheet(ByVal StaffName As String, ByVal StartDate As String, ByVal EndDate As String)
    
    
    Dim lastRowBudgetSheet As Integer
    Call AddStaffToWeekly(StaffName)

    Sheets("Budget").Unprotect
    
    lastRowBudgetSheet = Sheets("Budget").Range("E" & Rows.Count).End(xlUp).Row
    Sheets("Budget").Range("E" & lastRowBudgetSheet + 1).Value = StaffName 'The value passed through which can include '
    Dim FinalSheetName As String
    FinalSheetName = Left(Replace(StaffName, "'", ""), 30)
    
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = FinalSheetName 'The value passed through with the ' removed
    Sheets(FinalSheetName).Activate
    
    '########################################## Colouring tabs ##############################################
    
    With ActiveWorkbook.Sheets(FinalSheetName).Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.5
    End With
    
    Range("B2").Value = StaffName 'The value passed through which can include an '
    Range("B2:C3").MergeCells = True
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#######################################################################################################
    'Budget OF RATUL JINDAL CODE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#######################################################################################################
    
    Dim i As Integer
    Dim counter As Integer
    Dim nDays As Integer
    Dim ColDateBudget, ColDateEnd As Integer
    Dim templateRange As Range
    Dim tmpDate As Date
    Dim forRange As Range
    Dim lastCol As Integer
    Dim totalHours_rng As Range
    
    '''''##############Budgeting Column############'''''''''
    i = 5 ' Column F, next column
    ColDateBudget = i
    '''''''''''#################################'''''''''''''''''
    
    Set templateRange = Sheets("Data").Range("D1:E22")
    templateRange.Copy
    
    ActiveSheet.Range("B5").PasteSpecial
    ActiveSheet.Range("D5") = "Total Hours"
    
    Set totalHours_rng = Range("D7:D25")
    Range("D6").Formula = "=sum(" & totalHours_rng.Address & ")"
    
    '        ActiveSheet.Range("E6") = BudgetDate
    '        ActiveSheet.Range("E5") = Format(Cells(6, i - 1), "dddd")
    
    nDays = dateDiff("d", StartDate, EndDate)
    'Debug.Print "printing datediff........", nDays
    ' i is column
    For counter = 0 To nDays
        tmpDate = Format(DateValue(StartDate), "dd/mm/yyyy")
        'Debug.Print "new Budget date.............", BudgetDate
        ActiveSheet.Cells(6, i) = tmpDate + counter
        ActiveSheet.Cells(5, i) = Format(Cells(6, i), "dddd")
        
        If (ActiveSheet.Cells(5, i).Value = "Saturday" Or ActiveSheet.Cells(5, i).Value = "Sunday") Then
            ActiveSheet.Range(Cells(5, i), Cells(25, i)).Style = "Bad"
            ActiveSheet.Range(Cells(6, i), Cells(25, i)).Style = "Bad"
        Else
            ActiveSheet.Cells(5, i).Style = "20% - Accent1"
            ActiveSheet.Cells(5, i).Font.Bold = True
            ActiveSheet.Cells(6, i).Style = "20% - Accent1"
            ActiveSheet.Cells(6, i).Font.Bold = True
        End If
        i = i + 1
    Next counter
        
        
    If (i = 1) Then
        ColDateEnd = i
    Else
        ColDateEnd = i - 1
    End If
    
    
    
    For i = 7 To 25
        Set forRange = Range(Cells(i, 5), Cells(i, ColDateEnd))
        'Debug.Print forRange.AddressLocal
        Cells(i, 4).Formula = "=sum(" & forRange.Address & ")"
    Next i
        
    For i = 5 To ColDateEnd
        Cells(26, i).Formula = "=sum(" & Range(Cells(7, i), Cells(25, i)).Address & ")"
    Next i
            
    lastCol = ColDateEnd
    'Debug.Print "last col........", lastCol
    
    Call formatNewSheet(lastCol)
    
    'Debug.Print "printing last column,,,,,,,,,,,,,,,", lastCol
    '    Else
    '              MsgBox ("Sheet Already there dude")
    '        Exit Sub
    '
    '    End If '' end of sheet found

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#######################################################################################################
    'END OF RATUL JINDAL CODE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#######################################################################################################
                
                
End Sub




