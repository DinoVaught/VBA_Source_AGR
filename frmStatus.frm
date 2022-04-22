Option Explicit

Public Sub UpdateMain(sText As String)
    On Error Resume Next
    Me.txtMain.Text = sText
    UpdateProgressBar
    DoEvents
    
    
End Sub


Public Sub UpdateDetails(sText As String)
    Me.txtDetails.Text = sText
    ' Me.prgProgress.Value = Me.prgProgress.Value + 1
    UpdateProgressBar
    DoEvents
End Sub


Private Sub UpdateProgressBar()
    On Error Resume Next
    If Me.prgProgress.Value = Me.prgProgress.Max - 1 Then
        Me.prgProgress.Value = Me.prgProgress.Max
    Else
        Me.prgProgress.Value = Me.prgProgress.Value + 1
    End If
    
'    count = count + 1
'    Debug.Print count
    
End Sub


Private Sub Sleep(count As Integer)
    Dim i As Integer


    i = 0
    
    Do While i < count
        DoEvents
        i = i + 1
'        DoEvents
    Loop


End Sub


Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    Me.Caption = MessageCaption
    
    Me.txtDetails.Text = "Initializing"
    Me.txtMain.Text = "Initializing"
    
End Sub

'Private Sub UserForm_Terminate()
'  prgCounter = prgCounter
'End Sub
