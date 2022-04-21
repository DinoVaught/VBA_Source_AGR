Option Explicit

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Me.Caption = "(Secondary AGR) (Excel v" & Application.Version & ")  (" & Application.UserName & ") (" & Environ$("computername") & ")"
    Me.lblWeekNum = "Week " & Str(DatePart("ww", Format$(Now, "MM/DD/YYYY"), vbMonday, vbFirstJan1))
    Me.lblDate = Format$(Now, "MM/DD/YYYY")
    
End Sub
