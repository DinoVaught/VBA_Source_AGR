Option Explicit


Private selectedOldColor As Integer
Private selectedOldRange As Range

Private firstIteration As Boolean
Private disabled As Boolean


'' https://stackoverflow.com/questions/58282085/change-cell-color-when-it-selected-and-back-original-color-after-leaving-it
'' https://www.techrepublic.com/article/how-to-use-vba-to-change-the-active-cell-color-in-an-excel-sheet/
Public Sub SaveCellData(ByVal Target As Range)
    Dim currentColor As Integer

    If disabled = True Then Exit Sub

    currentColor = Target.Interior.ColorIndex
    
    Target.Interior.ColorIndex = 8 ' Cyan
    SetBorderColors Target
    
    
    If firstIteration = True Then
        firstIteration = False
        selectedOldColor = currentColor
        Set selectedOldRange = Target
        Exit Sub
    End If
    
    ResetPreviousCell
    selectedOldColor = currentColor
    Set selectedOldRange = Target
    
End Sub

Private Sub SetBorderColors(ByVal Target As Range)

    Debug.Print FindGage.GetGage(Target.Row)

End Sub

Public Sub ResetPreviousCell()

    If disabled = True Then Exit Sub
    
    
    If firstIteration = True Then Exit Sub
    selectedOldRange.Interior.ColorIndex = selectedOldColor
    
End Sub

Private Sub Class_Initialize()
    disabled = True
    firstIteration = True
End Sub

Private Sub Class_Terminate()
    ResetPreviousCell
End Sub
