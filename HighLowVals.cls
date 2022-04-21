Option Explicit

Private lowValue As Integer
Private highValue As Integer
Private station_1_Col As Integer

Public Property Get HighVal() As Integer
    HighVal = highValue
End Property

Public Property Let HighVal(ByVal val As Integer)
    highValue = val
End Property

Public Property Get LowVal() As Integer
    LowVal = lowValue
End Property

Public Property Let LowVal(ByVal val As Integer)
    lowValue = val
End Property

Public Property Get ST_1_Col() As Integer
    ST_1_Col = station_1_Col
End Property
