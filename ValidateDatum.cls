Option Explicit


Public Function EvalStationCount(resolvedValue As String, startValue As String) As String

    If Int(resolvedValue) = 0 And Int(startValue) = 0 Then
        EvalStationCount = resolvedValue
        Exit Function
    End If
    
    If Int(resolvedValue) < 0 Then
        EvalStationCount = "na"
        Exit Function
    End If

    If Int(startValue) = 0 Then
        EvalStationCount = "na"
        Exit Function
    End If
    
    EvalStationCount = resolvedValue
   
End Function


Public Function MassageDatum(datum As String) As String

    If Int(datum) < 0 Then
        MassageDatum = "na"
    Else
        MassageDatum = datum
    End If
   
End Function
