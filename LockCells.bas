Sub LockAllButAFew()
    
    'ActiveSheet.
    Worksheets("Budget").Protection.AllowEditRanges.Add Title:="Range1", Range:=Range( _
        "C5:C6,C19,C27:C30")
    'ActiveSheet.
    Worksheets("Budget").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

End Sub
Sub UnlockAll()

    'ActiveSheet.
    Worksheets("Budget").Unprotect
    'ActiveSheet.
    Worksheets("Budget").Protection.AllowEditRanges(1).Delete
End Sub

