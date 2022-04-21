Option Explicit

Private Const NO_DAY As String = "**"

Private Const MON As String = "Monday"
Private Const TUE As String = "Tuesday"
Private Const WED As String = "Wednesday"
Private Const THUR As String = "Thursday"
Private Const FRI As String = "Friday"
Private Const SAT As String = "Saturday"
Private Const SUN As String = "Sunday"


Private NextDay As String
Private EndOfWeek As Boolean


' this class returns values (weekday names) that match the names of (named ranges) in this spreadsheet

Public Property Get GetNextDay() As String
    
    ResolveNextDay
    GetNextDay = NextDay
    
End Property


Public Property Get EOW() As Boolean ' End Of Week
    EOW = EndOfWeek
End Property


Private Sub ResolveNextDay()

    Select Case NextDay
    
        Case NO_DAY
            NextDay = MON
        Case MON
            NextDay = TUE
        Case TUE
            NextDay = WED
        Case WED
            NextDay = THUR
        Case THUR
            NextDay = FRI
        Case FRI
            NextDay = SAT
        Case SAT
            NextDay = SUN
'        Case SUN
'            NextDay = vbNullString
'        Case Else
'            NextDay = vbNullString
        
    End Select
        
    EndOfWeek = NextDay = SUN

End Sub


Private Sub Class_Initialize()
    NextDay = NO_DAY
    EndOfWeek = False
End Sub
