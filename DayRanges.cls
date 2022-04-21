Option Explicit

Private Const MON As String = "Monday"
Private Const TUE As String = "Tuesday"
Private Const WED As String = "Wednesday"
Private Const THUR As String = "Thursday"
Private Const FRI As String = "Friday"
Private Const SAT As String = "Saturday"
Private Const SUN As String = "Sunday"
Private Const UNKNOWN_DAY As String = "NA"



Dim hlMon As New HighLowVals
Dim hlTue As New HighLowVals
Dim hlWed As New HighLowVals
Dim hlThur As New HighLowVals
Dim hlFri As New HighLowVals
Dim hlSat As New HighLowVals
Dim hlSun As New HighLowVals

Private dayName As String
Private selectedColumn As Integer


Public Property Get DayRangeName() As String
    DayRangeName = dayName
End Property

Public Property Let SelectedCol(ByVal val As Integer)
    selectedColumn = val
    ResolveDayName
End Property

Private Sub Class_Initialize()


    Dim currentDay As String
    Dim rangeLow As Integer
    Dim rangeHigh As Integer

    Dim DayNames As New DayNames
    
    Do Until DayNames.EOW = True

        currentDay = DayNames.GetNextDay()
        
        rangeLow = Range(currentDay).Column
        rangeHigh = Range(currentDay).Column + Range(currentDay).ColumnWidth + 1
        
        Select Case currentDay
        
            Case MON
                hlMon.LowVal = rangeLow
                hlMon.HighVal = rangeHigh
                
            Case TUE
                hlTue.LowVal = rangeLow
                hlTue.HighVal = rangeHigh
            
            Case WED
                hlWed.LowVal = rangeLow
                hlWed.HighVal = rangeHigh
            
            Case THUR
                hlThur.LowVal = rangeLow
                hlThur.HighVal = rangeHigh
            
            Case FRI
                hlFri.LowVal = rangeLow
                hlFri.HighVal = rangeHigh
            
            Case SAT
                hlSat.LowVal = rangeLow
                hlSat.HighVal = rangeHigh
                
            Case SUN
                hlSun.LowVal = rangeLow
                hlSun.HighVal = rangeHigh
                
        End Select


'        Application.GoTo Reference:=currentDay

        
    Loop

End Sub


Private Sub ResolveDayName()
    Select Case selectedColumn
        
        Case hlMon.LowVal To hlMon.HighVal
            dayName = MON
            
        Case hlTue.LowVal To hlTue.HighVal
            dayName = TUE
        
        Case hlWed.LowVal To hlWed.HighVal
            dayName = WED
        
        Case hlThur.LowVal To hlThur.HighVal
            dayName = THUR
            
        Case hlFri.LowVal To hlFri.HighVal
            dayName = FRI
            
        Case hlSat.LowVal To hlSat.HighVal
            dayName = SAT
            
        Case hlSun.LowVal To hlSun.HighVal
            dayName = SUN

        Case Else
            dayName = UNKNOWN_DAY
            
    End Select
End Sub
