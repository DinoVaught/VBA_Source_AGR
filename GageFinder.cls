Option Explicit

' Private Const BLANK_VAL As String = vbNullString
Private Const CYCLE_TIME As String = "CYCLE TIME"
Private Const SHFT_GOAL As String = "S GOAL"
Private Const WK_GOAL As String = "WK GOAL"


Public Function GetGage(rowNum As Integer) As String
    
    Dim trgtVal As String
    
    trgtVal = UCase$(Cells((rowNum), 1).Value)
    
    If trgtVal = CYCLE_TIME Then GetGage = ResolveGageName(CYCLE_TIME, rowNum): Exit Function

    If trgtVal = SHFT_GOAL Then GetGage = ResolveGageName(SHFT_GOAL, rowNum): Exit Function
    
    If trgtVal = WK_GOAL Then GetGage = ResolveGageName(WK_GOAL, rowNum): Exit Function
    
    trgtVal = UCase$(Cells((rowNum - 1), 1).Value)
    If trgtVal = CYCLE_TIME Then GetGage = ResolveGageName(CYCLE_TIME, rowNum - 1): Exit Function
    
    trgtVal = UCase$(Cells((rowNum - 2), 1).Value)
    If trgtVal = CYCLE_TIME Then GetGage = ResolveGageName(CYCLE_TIME, rowNum - 2): Exit Function
    
    trgtVal = UCase$(Cells((rowNum + 1), 1).Value)
    If trgtVal = CYCLE_TIME Then GetGage = ResolveGageName(CYCLE_TIME, rowNum + 1): Exit Function
    
    
    
End Function


Private Function ResolveGageName(tag As String, rowNum As Integer) As String

    
    Dim gageRow As Integer
    
    Select Case tag
    
        Case CYCLE_TIME
            gageRow = rowNum
            
        Case SHFT_GOAL
            gageRow = rowNum - 3
            
        Case WK_GOAL
            gageRow = rowNum - 4
            
    End Select

    ResolveGageName = Cells(gageRow, 2).Value

End Function
