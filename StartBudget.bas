'ADD STAFF BUTTON ON THE Budget SHEET.
Private Sub BtnAddStaff_Click()

'    LstStaffMembers.ColumnHeads = False
'    LstStaffMembers.RowSource = "Data!Z2:AA2"
'    Sheets("Data").Range("Z1:AA1").Address(external:=True)
    
    

    Dim i, j As Integer
    Dim item As String
    Dim validEntry, itemFound As Boolean
    
    validEntry = False
    itemFound = False
    
    i = LstStaffMembers.ListCount
    For j = 0 To i - 1
        item = LstStaffMembers.List(j)
        
        If (item = CmbStaffName.Text) Then
            itemFound = True
    '        MsgBox ("Staff Already Exists!!")
            Exit For
'        ElseIf (item = PartnerComboBox.Text) Then
'            LstStaffMembers.RemoveItem (j)
        Else
            itemFound = False
        End If
    Next j
    
        If (CmbStaffName.Text = " " And CmbGrade.Text = " ") Then
            validEntry = False
        ElseIf (CmbGrade.Text = "" Or CmbStaffName.Text = "") Then
            validEntry = False
        ElseIf (itemFound = True) Then
            validEntry = False
        ElseIf (PartnerComboBox.Text = CmbStaffName.Text) Then
            validEntry = False
        ElseIf (ManagerComboBox.Text = CmbStaffName.Text) Then
            validEntry = False
        Else
            validEntry = True
        End If
'
'    With LstStaffMembers
'        .ColumnCount = 2
'        .ColumnWidths = "170;30"
'        .AddItem
'        .List(0, 0) = "Staff Members"
'        .List(0, 1) = "Staff Grade"
'    End With

'    Dim i As Integer

'    i = 0
'    For i = 0 To LstStaffMembers.ListCount
'        If PartnerComboBox.Text = LstStaffMembers.List(i) Then
'            LstStaffMembers.RemoveItem (i)
'            Exit For
'        End If
'
'
'
'    Next i
'
'
     i = LstStaffMembers.ListCount
    
    If validEntry = True Then
         With LstStaffMembers
            .ColumnCount = 2
            .ColumnWidths = "170;30"
            .AddItem
            .List(i, 0) = CmbStaffName.Text
            .List(i, 1) = CmbGrade.Text
            i = i + 1
        End With
    Else
        MsgBox ("Staff Member Already Exists!")
    
    End If
    
'    LstStaffMembers.sor
    
    'LstStaffMembers.AddItem (CmbStaffName.Text)

End Sub

Private Sub BtnDeleteStaff_Click()
 If LstStaffMembers.ListCount = 0 Then
        
    ElseIf LstStaffMembers.ListIndex = -1 Then
'    Or LstStaffMembers.ListIndex = 0 Or LstStaffMembers.ListIndex = 1 Then
'    MsgBox ("Project Partner and Manager cannot be removed, Please check your listbox" & vbNewLine & "Or, you need to Add at least a Staff member first")
        MsgBox ("Please add at least 1 Staff Member to the list")
    Else
        LstStaffMembers.RemoveItem (LstStaffMembers.ListIndex)
    End If
End Sub

'CLEARS THE STAFF NAME LIST WHEN SELECTION ON THE GRADE IS CHANGED.
Private Sub CmbGrade_Change()
CmbStaffName.Clear
End Sub

'SEARCH FUNCTION FOR PE CODE
Private Sub CmbPECode_Change()
Dim ws As Worksheet
Dim x, dict
Dim i As Long
Dim str As String

Set ws = Sheets("Client_Codes")
x = ws.Range("Client_Code").Value

Set dict = CreateObject("Scripting.Dictionary")
str = Me.CmbPECode.Value

If str <> "" Then
    For i = 1 To UBound(x, 1)
        If InStr(LCase(x(i, 1)), LCase(str)) > 0 Then
            dict.item(x(i, 1)) = ""
        End If
    Next i
    Me.CmbPECode.List = dict.keys
Else
    Me.CmbPECode.List = x
End If
Me.CmbPECode.DropDown

FrameResources.Visible = False
FrameDate.Visible = False
FrameFinal.Visible = False

'Call CommandButton1_Click



End Sub

Private Sub CmbPECode_Click()
'    CommandButton1.Enabled = True
End Sub

Private Sub CmbStaffName_DropButtonClick()

Dim gradeType As String

gradeType = CmbGrade.Text
'Debug.Print gradeType

Dim cLoc As Range
Dim ws As Worksheet
Dim i As Integer

Set ws = Worksheets("Data")
'CmbStaffName.Clear
    For Each cLoc In ws.Range(gradeType)
        With Me.CmbStaffName
             .AddItem cLoc.Value
        End With
    Next cLoc
    
i = LstStaffMembers.ListCount
  For j = 0 To i - 1
      item = LstStaffMembers.List(j)
   
   If (item = PartnerComboBox.Text) Then
          LstStaffMembers.RemoveItem (j)
    ElseIf (item = ManagerComboBox.Text) Then
        LstStaffMembers.RemoveItem (j)
      Else
'          itemFound = False
      End If
  Next j



End Sub


 
'PE CODE SEARCH BUTTON ON THE FORM WHICH WILL POPULATE THE PARTNER AND MANAGER NAME
'THE PARTNER AND MANAGER NAME IS COMING FROM THE DATA SHEET.
Private Sub CommandButton1_Click()
'MsgBox CmbPECode.Text

    Dim PE_Code As String
    Dim partnerName As String
    Dim managerName As String
    Dim clientName As String
    Dim rowEnd, i As Integer
    Dim PECodeFound As Boolean
    
    PECodeFound = False
    
    rowEnd = Sheets("Client_Codes").Range("A" & Rows.Count).End(xlUp).Row
    
    LstStaffMembers.Clear
    PartnerComboBox.Clear
    ManagerComboBox.Clear
    
    PE_Code = CmbPECode.Text
    
    For i = 2 To rowEnd
        If Sheets("Client_Codes").Range("A" & i).Value <> PE_Code Then
            
        Else
            PECodeFound = True
            clientName = Sheets("Client_Codes").Range("B" & i).Value
            PartnerComboBox.AddItem (Sheets("Client_Codes").Range("C" & i).Value)
            ManagerComboBox.AddItem (Sheets("Client_Codes").Range("D" & i).Value)
            Exit For
        End If
        
    Next i
    
    
'    If (PECodeFound = True) Then
    
    
    
    

'    clientName = Application.WorksheetFunction.IfError(Application.WorksheetFunction _
'    .VLookup(PE_Code, Sheets("Client_Codes").Range("A:D"), 2, False), "")


    
    
    If (PECodeFound = False) Then
        MsgBox ("Please select a valid Client Code from the list")
        ClientNameTxtBox.Text = ""
    Else
        
        
    ClientNameTxtBox.Text = clientName
    
        If (ClientNameTxtBox.Text = "" Or PE_Code = "") Then
    '        MsgBox "Select PE Code first"
        Else
        
            FrameResources.Visible = True
            FrameDate.Visible = True
            FrameFinal.Visible = True
            
        FinalStartDatePicker.Value = Date
        FinalEndDatePicker.Value = Date
        
        FeeTextBox.Value = Int(0)
        
        Dim cLoc As Range
        Dim ws As Worksheet
        
        Set ws = Worksheets("Data")
        'CmbStaffName.Clear
        
'
'        PartnerComboBox.AddItem (WorksheetFunction.VLookup(PE_Code, Sheets("Client_Codes").Range("A:D"), 3, False))
'        ManagerComboBox.AddItem (WorksheetFunction.VLookup(PE_Code, Sheets("Client_Codes").Range("A:D"), 4, False))
        
            For Each cLoc In ws.Range("Audit_Partners")
                With Me.PartnerComboBox
                     .AddItem cLoc.Value
                End With
            Next cLoc
            
            For Each cLoc In ws.Range("dsma_group")
                With Me.ManagerComboBox
                    If (cLoc.Value <> "") Then
                    
                     .AddItem cLoc.Value
                    Else
                    End If
                    
                End With
            Next cLoc
            
        
        End If
        
        PartnerComboBox.Text = PartnerComboBox.List(0)
        ManagerComboBox.Text = ManagerComboBox.List(0)
    
    End If

End Sub



Private Sub CommandButton2_Click()

LstStaffMembers.Clear

End Sub

'COMMA SEPARATE VALUES FOR FEES TEXTBOX
Private Sub FeeTextBox_Change()
'FeeTextBox.
    FeeTextBox.Value = Format(FeeTextBox.Value, "#,##0")

End Sub

Private Sub FinalStartDatePicker_DateClick(ByVal DateClicked As Date)
    'Ensures that the date selected is further into the future than todays date and also is a weekday
'    If Weekday(DateClicked, vbMonday) < 6 Then
        'If DateClicked < Date Then
            'MsgBox "Future Date Must be selected"
        'Else
            FinalStartDateTextBox.Text = CStr(DateClicked)
            'If FinalEndDateTextBox.Text <> "" Then

            'End If
       ' End If
    'Else
        'MsgBox "Please Select Weekdays Only"
'    End If
End Sub


Private Sub FinalEndDatePicker_DateClick(ByVal DateClicked As Date)
    'Ensures that the date selected is further into the future than todays date and also is a weekday
'    If Weekday(DateClicked, vbMonday) < 6 Then
'        If DateClicked < Date Then
'            MsgBox "Future Date Must be selected"
'        Else
            FinalEndDateTextBox.Text = CStr(DateClicked)
'            If FinalBudgetDateTextBox.Text <> "" Then
'
'            End If
'        End If
'    Else
'        MsgBox "Please Select Weekdays Only"
'    End If
End Sub

'SHEETS ARE CREATED WHEN THE SUBMIT BUTTON IS CLICKED AND SUMMARY AND WEEKLY ARE POPULATED.
Private Sub SubmitStartData_Click()

fee.btnFeeBreakCounter = 1

Application.ScreenUpdating = False

'        StartBudget.Hide
'This function populates the data in the Budget tab in the correct positions for
'the rest of the features to work correctly

    Dim PE_Code, Client_Name, Partner_Name, Manager_Name, Budget_Date, End_Date, _
        FInvoice, IInvoice As String
        
    Dim staffCount As Integer
    Dim validDate As Boolean
    Dim auditFee As Double
    
    validDate = False
    
    
    PE_Code = CmbPECode.Text
    Client_Name = ClientNameTxtBox.Text
    Partner_Name = PartnerComboBox.Text
    Manager_Name = ManagerComboBox.Text
    
'    MsgBox (auditFee)
'
    staffCount = LstStaffMembers.ListCount
    
    FStart_Date = FinalStartDateTextBox.Text
    FEnd_Date = FinalEndDateTextBox.Text
    
    If (FStart_Date = "" Or FEnd_Date = "") Then
        validDate = False
    ElseIf CDate(FStart_Date) > CDate(FEnd_Date) Then
        validDate = False
    Else
        validDate = True
    End If
    
    If validDate = False Then
        MsgBox "Final Budget Date cannot be after Budget End Date"
    ElseIf Client_Name = "" Then
        MsgBox "Please Enter A Client"
    ElseIf PE_Code = "" Then
        MsgBox "Please Enter A PE Code"
    ElseIf Partner_Name = "" Then
        MsgBox "Please Select A Partner"
    ElseIf Manager_Name = "" Then
        MsgBox "Please Select A Manager"
    ElseIf IsNumeric(FeeTextBox.Text) = False Then
        MsgBox "Please Enter a Numeric value"
        FeeTextBox.Text = 0
    ElseIf IsNumeric(FeeTextBox.Text) <> True Then
        MsgBox "Please Enter a Numeric Value"
    ElseIf staffCount = 0 Then
        MsgBox ("Please select at least 1 Team Member")
    
    Else
        Sheets("Budget").Unprotect
        Call WeeklySummarySetup(FStart_Date, FEnd_Date)
        
        auditFee = Int(FeeTextBox.Text)
        With Sheets("Budget")
            .Range("C8").Value = PE_Code
            .Range("C8").HorizontalAlignment = xlRight
            
            .Range("C9").Value = Client_Name
            .Range("C9").HorizontalAlignment = xlRight
            
            .Range("C11").Value = Partner_Name
            .Range("C11").HorizontalAlignment = xlRight
            
            .Range("C12").Value = Manager_Name
            .Range("C12").HorizontalAlignment = xlRight
 
            .Range("C16").Value = CDate(FStart_Date)
            .Range("C16").HorizontalAlignment = xlRight
        
            .Range("C17").Value = CDate(FEnd_Date)
            .Range("C17").HorizontalAlignment = xlRight
        
        
            .Range("C19").Value = auditFee
            .Range("C19").Style = "Currency"
            .Rows("9:9").EntireRow.AutoFit
        End With
        
        StartDate = CStr(Worksheets("Budget").Range("C16").Value)
        EndDate = CStr(Worksheets("Budget").Range("C17").Value)
      
                
        '############################################# Adds The Project Staff ##########################################
          
        Call CreateNewSheet(Partner_Name, StartDate, EndDate)
        Call CreateNewSheet(Manager_Name, StartDate, EndDate)
        
        i = LstStaffMembers.ListCount
        
        For j = 0 To i - 1
            item = LstStaffMembers.List(j)
'            Sheets("Budget").Cells(6 + j, "e").Value = item
            Call CreateNewSheet(item, StartDate, EndDate)
        Next j
        
      
        Sheets("Budget").Protect
'        Sheets("Budget").Select
        
        
        
        Call summarySheet
        Call weeklySum
        Call feeBreakDown
        'Call WeeklySummarySetup(FStart_Date, FEnd_Date)
        
        
        'BtnAddStaff.Enabled = True
        
        StartBudget.Hide
    End If
    
    
    
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Activate()

'BudgetBudget.Show
  
End Sub

'POPULATES THE STAFF GRADE WHEN THE FORM IS INITIALIZED

Private Sub UserForm_Initialize()

    
    FrameResources.Visible = False
    FrameDate.Visible = False
    FrameFinal.Visible = False
    
 
    Dim cLoc As Range
    Dim ws As Worksheet
    Set ws = Worksheets("Data")
    
    For Each cLoc In ws.Range("GradesList")
        With Me.CmbGrade
             .AddItem cLoc.Value
        End With
    Next cLoc
'    CmbPECode.Text
End Sub