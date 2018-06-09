Private Sub Worksheet_Activate()
     Sheets("Summary").Unprotect
        Sheets("Summary").PivotTables("AuditPivotTable").RefreshTable
        Sheets("Summary").Protect
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub