Option Explicit


Public Function MassageDatum(datum As String) As String

    Dim datValue As String
    
    If Int(datum) < 0 Then
        MassageDatum = "na"
    Else
        MassageDatum = datum
    End If
   
End Function
