Private Sub BtnAddStaffList_Click()
    Dim i As Integer
    Dim StaffName, TempStr As String
    Dim wks As Variant
    Dim staffFound, AddedToList As Boolean
    AddedToList = False
    staffFound = False
    StaffName = CmbSelectStaff.Text
    
    i = LstBoxAddStaff.ListCount
    
    For Each wks In Worksheets
        If (wks.Name = StaffName) Then
            staffFound = True
            Exit For
        Else
            
        End If
    Next wks
    
    
    For a = 0 To i - 1
        TempStr = LstBoxAddStaff.List(a, 0)
        If TempStr = StaffName Then
            AddedToList = True
            'MsgBox "The Staff Member " & StaffName & " is already in your list to add"
        End If
    Next a
    
    
    
    If StaffName = "" Then
    
    ElseIf AddedToList = True Then
            
    ElseIf staffFound = False Then
        With LstBoxAddStaff
                .ColumnCount = 2
                .ColumnWidths = "100;30"
                .AddItem
                .List(i, 0) = CmbSelectStaff.Text
                .List(i, 1) = CmbSelectGrade.Text
        '        i = i + 1
        End With
    Else
        
        MsgBox (StaffName & " already in the Project")
    
    End If
End Sub




Private Sub BtnAddTrainee_Click()
    Dim trainee As String
    Dim listCounter As Integer
    
    listCounter = LstBoxAddStaff.ListCount
    
    trainee = InputBox("How many Trainees do you want to add?", "Add Trainee")
    
    trainee = Int(trainee)
    
    Dim i As Integer
    
    For i = 1 To trainee
        With LstBoxAddStaff
                .ColumnCount = 2
                .ColumnWidths = "100;30"
                .AddItem
                .List(listCounter, 0) = "Trainee_" & str(i)
                .List(listCounter, 1) = "Trainee"
            listCounter = listCounter + 1
        End With
    Next i
End Sub




Private Sub BtnDeleteStaff_Click()
    If LstBoxAddStaff.ListCount = 0 Then
        
    ElseIf LstBoxAddStaff.ListIndex = -1 Then
    
    Else
        LstBoxAddStaff.RemoveItem (LstBoxAddStaff.ListIndex)
    End If
    
End Sub




Private Sub BtnDone_Click()
    
    Sheets("Budget").Unprotect
    Dim SDate, EDate As String
    SDate = Worksheets("Budget").Range("C16").Text
    EDate = Worksheets("Budget").Range("C17").Text
    
    
'=============================| Adds the people from the list box first |===================
    If LstBoxAddStaff.ListCount > 0 Then
        Dim StaffName, TempStaffName, CellValue As String
        Dim CountOfEntries, EndRow, InStrPos As Integer
        Dim WasFound As Boolean
        CountOfEntries = 0
        Range("E5").Select
        EndRow = Selection.End(xlDown).Row
        
        '+-----------------------------------------------------+
        '|  Checks to see if the Person's name already exists  |
        '+-----------------------------------------------------+
        For i = 0 To LstBoxAddStaff.ListCount - 1
            staffMember = LstBoxAddStaff.List(i)
            WasFound = False
            For a = 5 To EndRow
                CellValue = Cells(a, 5).Text
                InStrPos = InStr(1, CellValue, staffMember)
                If InStrPos = 0 Then
                    '+-------------------------------+
                    '|  If their name was not found  |
                    '+-------------------------------+
                    'Call CreateNewSheet(staffMember, SDate, EDate)
                ElseIf InStrPos > 0 Then
                    '+---------------------------+
                    '|  If their name was found  |
                    '+---------------------------+
                    WasFound = True
                    'MsgBox staffMember & " already exists!!"
                End If
            Next a
            If WasFound = False Then
                Call CreateNewSheet(staffMember, SDate, EDate)
            Else
                MsgBox staffMember & " already exists!!"
            End If
        Next i
    
        Sheets("Budget").Protect
        Call summarySheet
        Call weeklySum
        Call feeBreakDown
    End If
'===========================================================================================
    
'=============================| Adds generic accounts second |==============================
    Dim CoOpCount, T1Count, T2Count, T3Count, SeniorCount, AMCount, TaxCount, RACount, ActuarialCount, ValuationCount As String

    CoOpCount = CoOpTextBox.Value
    T1Count = T1TextBox.Value
    T2Count = T2TextBox.Value
    T3Count = T3TextBox.Value
    SeniorCount = SeniorTextBox.Value
    AMCount = AMTextBox.Value
    TaxCount = TaxTextBox.Value
    RACount = RATextBox.Value
    ActuarialCount = ActuarialTextBox.Value
    ValuationCount = ValuationTextBox.Value
    
    
    If CoOpCount <> "" And IsNumeric(CoOpCount) Then
        Call loopStaffAdd(CInt(CoOpCount), "Co-Op", SDate, EDate)
    End If
    
    If T1Count <> "" And IsNumeric(T1Count) Then
        Call loopStaffAdd(CInt(T1Count), "Trainee 1", SDate, EDate)
    End If
    
    If T2Count <> "" And IsNumeric(T2Count) Then
        Call loopStaffAdd(CInt(T2Count), "Trainee 2", SDate, EDate)
     End If
     
    If T3Count <> "" And IsNumeric(T3Count) Then
        Call loopStaffAdd(CInt(T3Count), "Trainee 3", SDate, EDate)
    End If
    
    If SeniorCount <> "" And IsNumeric(SeniorCount) Then
        Call loopStaffAdd(CInt(SeniorCount), "Senior", SDate, EDate)
    End If
    
    If AMCount <> "" And IsNumeric(AMCount) Then
        Call loopStaffAdd(CInt(AMCount), "Assistant Manager", SDate, EDate)
    End If
    
    If TaxCount <> "" And IsNumeric(TaxCount) Then
        Call loopStaffAdd(CInt(TaxCount), "Tax Specialist", SDate, EDate)
    End If
    
    If RACount <> "" And IsNumeric(RACount) Then
        Call loopStaffAdd(CInt(RACount), "RA Specialist", SDate, EDate)
    End If
    
    If ActuarialCount <> "" And IsNumeric(ActuarialCount) Then
        Call loopStaffAdd(CInt(ActuarialCount), "Actuarial Specialist", SDate, EDate)
    End If
    
    If ValuationCount <> "" And IsNumeric(ValuationCount) Then
        Call loopStaffAdd(CInt(ValuationCount), "Valuation Specialist", SDate, EDate)
    End If
  
    Call summarySheet
    Call weeklySum
    Call feeBreakDown
    
   
    FrmSelectStaff.Hide
    
End Sub
    
Sub loopStaffAdd(ByVal Count As Integer, ByVal TypeOfStaff As String, ByVal BudgetDate As String, ByVal EndDate As String)
    Dim StaffName, TempStaffName, CellValue As String
    Dim CountOfEntries, EndRow, InStrPos As Integer
    StaffName = TypeOfStaff & "_"
    CountOfEntries = 0
    Range("E5").Select
    EndRow = Selection.End(xlDown).Row
    
    
    
    
    
    For a = 5 To EndRow
        CellValue = Cells(a, 5).Text
        InStrPos = InStr(1, CellValue, StaffName)
        If InStrPos = 0 Then
            'MsgBox "Not Found"
        ElseIf InStrPos > 0 Then
            'MsgBox "Found"
            CountOfEntries = CountOfEntries + 1
        End If
    Next a
    
    
    If Count = 0 Then
        'MsgBox "Worked"
    ElseIf (CountOfEntries + Count) > 5 Then
        MsgBox "You cannot have more than 5 generic placeholders for " & TypeOfStaff & vbCrLf & vbCrLf & _
        "The operation to add " & Count & " " & TypeOfStaff & " has been cancelled!!"
    ElseIf Count > 0 And Count <= 5 Then
        For i = 1 To Count
            
            TempStaffName = StaffName & (CountOfEntries + i)
            Application.ScreenUpdating = False
            Call CreateNewSheet(TempStaffName, BudgetDate, EndDate)
        Next i
    End If
    
    
End Sub



Private Sub CmbSelectGrade_Change()
    CmbSelectStaff.Clear
    Dim gradeType As String
    
    gradeType = CmbSelectGrade.Text
    
    Dim cLoc As Range
    Dim ws As Worksheet
    
    Set ws = Worksheets("Data")
    
    For Each cLoc In ws.Range(gradeType)
        With Me.CmbSelectStaff
             .AddItem cLoc.Value
        End With
    Next cLoc
End Sub




Private Sub UserForm_Initialize()
   
    Dim cLoc As Range
    Dim ws As Worksheet
    
    Set ws = Worksheets("Data")
    
    For Each cLoc In ws.Range("GradesList")
        With Me.CmbSelectGrade
             .AddItem cLoc.Value
        End With
    Next cLoc

End Sub
