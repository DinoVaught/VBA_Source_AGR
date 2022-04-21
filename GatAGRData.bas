Option Explicit


Public Sub GetDataMac() ' For user keypress <ctrl><r>, Needs to be a Sub and not a Function

    
    On Error GoTo GetDataMacErrHandler
    
    InitializeEnvironment
    
    Application.Interactive = False
    frmStatus.UpdateMain ActiveCell.Text
    frmStatus.UpdateDetails "Getting Data"
    frmStatus.Show vbModeless
    frmStatus.prgProgress.Max = 81
    DoEvents
    If GetData = False Then
        ActivateUI
        End
    End If
    
    CheckMissingDataMsg
    
GetDataMacExitPoint:
    On Error Resume Next
    Application.Interactive = True
    ActivateUI
    Exit Sub
    
GetDataMacErrHandler:
    
    MsgBox Err.Description, vbInformation, MessageCaption
    
    Resume GetDataMacExitPoint
    
    
End Sub

Public Function GetData() As Boolean

    
    Const ODBC_MISSING As Long = -2147467259
    ' get week from year (below) where 11 is the week number
    ' DateAdd("ww", 11, DateSerial(Year(Date), 1, -4))
    
    ' Application.ActiveWorkbook.Path
    
    Dim inDate As Date
    Dim rsAGRData As New ADODB.Recordset
        
    On Error GoTo GetDataErrHandler
    GetData = True

    Application.Cursor = xlWait
    
    ValidateEnvironment
    
    If DateValid = False Then
        Exit Function ' error msg already displayed in DateValid function
    End If
    
    inDate = Range("D1").Value
    
    inDate = DateAdd("d", GetDayInt(ActiveCell.Value) - 2, inDate)
    ' inDate = DateAdd("d", GetDayInt(DayRange.DayRangeName) - 2, inDate)
    
    
    Set rsAGRData = QueryData(inDate)
    PopulateSheet rsAGRData
    

    ' Debug.Print (DateAdd("ww", 1 - 1, inDate))
    
        
GetDataExitPoint:
    On Error Resume Next
    Application.Cursor = xlDefault
    ActivateUI
    Exit Function
    
GetDataErrHandler:
    
    Select Case Err.Number
    
    
        Case ODBC_MISSING
            Unload frmStatus
            MsgBox "Cannot connect to the mes database" & vbCrLf & vbCrLf & "ODBC Data source is missing or configured incorrectly", vbExclamation, MessageCaption()
            
        Case Else
            MsgBox "Error occurred in GetData" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, MessageCaption
        
    End Select
        
    
    
    GetData = False
    Resume GetDataExitPoint

End Function

Public Function GetDataAllDays(sheetName As String) As Boolean
    
    Dim DayNames As New DayNames
    Dim currentDay As String
    
    On Error GoTo GetDataAllDaysErrHandler
    
    GetDataAllDays = False
    
    Sheets(sheetName).Select
    
    Do Until DayNames.EOW = True
    
     
        ' Debug.Print DayNames.GetNextDay()
        currentDay = DayNames.GetNextDay()
        frmStatus.UpdateMain currentDay
        DoEvents
        Application.GoTo Reference:=currentDay
        If GetData = False Then
            Exit Function
        End If
        
    Loop
    
    GetDataAllDays = True
    

GetDataAllDaysExitPoint:
    Exit Function
GetDataAllDaysErrHandler:
    GetDataAllDays = False
End Function



Private Sub PopulateSheet(rsData As ADODB.Recordset)

    Dim clsAGR As New AGR_Range
    Dim dataValid As New ValidateDatum
    Dim targetRow As Integer
    Dim errData As String
    
    ClearData
    Set rsData = DisconnectRecordSet(rsData)
    
    If rsData.EOF And rsData.BOF Then
        Exit Sub
    End If
    
    
    rsData.MoveFirst

    Dim cell As Range

    Do Until rsData.EOF = True
        
       
        frmStatus.UpdateDetails rsData!Gage_ID
        
       
        clsAGR.Initialize rsData!Gage_ID, rsData!partNum, ActiveCell.Column + 1
        
        If clsAGR.ErrorMsg = vbNullString Then
        
            Select Case rsData!shift
                Case "3"
                    targetRow = clsAGR.Shift3Row
                    
                Case "1"
                    targetRow = clsAGR.Shift1Row
                
                Case "2"
                    targetRow = clsAGR.Shift2Row
    
            End Select
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_1)) = False Then
                Cells(targetRow, clsAGR.ST_1_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_1_Col
            Else
                Cells(targetRow, clsAGR.ST_1_Col).Value = dataValid.MassageDatum(rsData!ST_1)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_2)) = False Then
                Cells(targetRow, clsAGR.ST_2_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_2_Col
            Else
                Cells(targetRow, clsAGR.ST_2_Col).Value = dataValid.MassageDatum(rsData!ST_2)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_3)) = False Then
                Cells(targetRow, clsAGR.ST_3_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_3_Col
            Else
                Cells(targetRow, clsAGR.ST_3_Col).Value = dataValid.MassageDatum(rsData!ST_3)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_4)) = False Then
                Cells(targetRow, clsAGR.ST_4_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_4_Col
            Else
                Cells(targetRow, clsAGR.ST_4_Col).Value = dataValid.MassageDatum(rsData!ST_4)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_5)) = False Then
                Cells(targetRow, clsAGR.ST_5_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_5_Col
            Else
                Cells(targetRow, clsAGR.ST_5_Col).Value = dataValid.MassageDatum(rsData!ST_5)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!ST_6)) = False Then
                Cells(targetRow, clsAGR.ST_6_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.ST_6_Col
            Else
                Cells(targetRow, clsAGR.ST_6_Col).Value = dataValid.MassageDatum(rsData!ST_6)
            End If
            
            If IsNumeric(dataValid.MassageDatum(rsData!Total)) = False Then
                Cells(targetRow, clsAGR.Total_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.Total_Col
            Else
                Cells(targetRow, clsAGR.Total_Col).Value = dataValid.MassageDatum(rsData!Total)
            End If
            
            
            If IsNumeric(dataValid.MassageDatum(rsData!AGR)) = False Then
                Cells(targetRow, clsAGR.AGR_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.AGR_Col
            Else
                Cells(targetRow, clsAGR.AGR_Col).Value = dataValid.MassageDatum(rsData!AGR)
            End If

            
            If IsNumeric(dataValid.MassageDatum(rsData!Net)) = False Then
                Cells(targetRow, clsAGR.NET_Col).Value = "0"
                Formatting.FormatCellOffsetColor targetRow, clsAGR.NET_Col
            Else
                Cells(targetRow, clsAGR.NET_Col).Value = dataValid.MassageDatum(rsData!Net)
            End If
            
        End If
    
        rsData.MoveNext
    Loop
    
    
End Sub

'Private Sub MassageData(ByRef rsData As ADODB.Recordset)
'
'    On Error GoTo MassageDataErrHandler
'
'    Set rsData = DisconnectRecordSet(rsData)
'
'    rsData.MoveFirst
'
'    Do Until rsData.EOF = True
'
'        If Int(rsData!ST_1) < 0 Then
'
'        End If
'
'        rsData.MoveNext
'    Loop
'
'
'MassageDataExitPoint:
'    On Error Resume Next
'    Exit Sub
'
'MassageDataErrHandler:
'
'    MsgBox "Error occurred in MassageData" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "Error"
'    Resume MassageDataExitPoint
'
'End Sub



Private Sub ClearData()  ' This sub assumes a Day (like Monday, Tuesday etc.. .) is the Active Cell / selected on the spreadsheet
    Dim GageCell As Range
    Dim firstRow As Integer
    Dim colAddress As String
    
    colAddress = Columns(ActiveCell.Column + 1).Address
    
    
    Set GageCell = Sheets(ActiveSheet.Index).Range(colAddress).Find("ST_1")
    firstRow = GageCell.Row
    
    ClearBlock GageCell
    
    Do
        ' Debug.Print GageCell.row
        
        Set GageCell = Sheets(ActiveSheet.Index).Range(colAddress).FindNext(GageCell)
        If firstRow <> GageCell.Row Then
            ClearBlock GageCell
        End If
        
        
        
    Loop Until firstRow = GageCell.Row
    
End Sub

Private Sub ClearBlock(startCell As Range)

    Dim crntRow As Integer
    Dim col As Integer
    
    crntRow = startCell.Row + 1
    
    
    Do Until crntRow = startCell.Row + 4
        
        col = startCell.Column
        
        Do Until col > startCell.Column + 8
            Cells(crntRow, col).Value = vbNullString
            Formatting.FormatCellNone crntRow, col
            col = col + 1
        Loop
    
    
        crntRow = crntRow + 1
    Loop
    
    
    

End Sub

Private Function QueryData(inDate As Date) As ADODB.Recordset

    Dim cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rsResults As New ADODB.Recordset
    Dim prmDate As ADODB.Parameter
    ' Dim row As String

    cnn.ConnectionString = "DSN=MySQL_Mes;uid=mes;pwd=mykiss;database=mes"
    cnn.Open
    cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GetAGR_Data_By_Day"
    
    rsResults.CursorLocation = adUseClient
    rsResults.LockType = adLockBatchOptimistic
    
    
    Set prmDate = cmd.CreateParameter("qryDate", adVarChar, adParamInput, 20)

    Dim sDate As String
    sDate = Year(inDate) & "-" & Month(inDate) & "-" & Day(inDate) ' format ("2022-02-27" "yyyy-mm-dd") or ("2022-2-7" "yyyy-m-d")
    prmDate.Value = Trim$(sDate)
    cmd.Parameters.Append prmDate
    rsResults.Open cmd
    Set QueryData = rsResults

End Function


Private Function DateValid() As Boolean
    On Error GoTo IsDateErrHandler
    
    Dim trgtDate As Date
    
    
    If IsDate(Range("D1").Value) = False Then
        Range("D1:H2").Select
        MsgBox "Cell D1 = " & Range("D1").Value & vbCrLf & vbCrLf & "Invalid date!" & vbCrLf & vbCrLf & "Please enter a valid date in Cell D1", vbOKOnly, MessageCaption
        DateValid = False
        Exit Function
    End If
    
    If (Weekday(Range("D1").Value) = vbMonday) = False Then
    
        Range("D1:H2").Select
        MsgBox "Cell D1 = " & Range("D1").Value & vbCrLf & vbCrLf & "Invalid date!" & vbCrLf & vbCrLf & "Please enter the (Monday) date for the corresponding week", vbOKOnly, MessageCaption
        DateValid = False
        Exit Function
    
    End If
    
    trgtDate = Range("D1").Value
    
    
    DateValid = True
    Exit Function
    
IsDateErrHandler:
    
    DateValid = False
    MsgBox "invalid date!"
    
End Function




Private Function GetDayInt(sDay As String) As Integer
    
    Select Case sDay
    
        Case "Monday"
            GetDayInt = vbMonday
        
        Case "Tuesday"
            GetDayInt = vbTuesday
            
        Case "Wednesday"
            GetDayInt = vbWednesday
        
        Case "Thursday"
            GetDayInt = vbThursday
        
        Case "Friday"
            GetDayInt = vbFriday
        
        Case "Saturday"
            GetDayInt = vbSaturday
        
        Case "Sunday"
            GetDayInt = 8 ' vbSunday = 1, vbSunday makes the date calc (for sunday) go backwards 1 day
                          ' this is why use an 8 here
        Case Else
            MsgBox "(Please select a day (Monday, Tuesday etc) from top part of the spreadsheet", vbInformation, MessageCaption
            ActivateUI
            End

    End Select
    
    
End Function

Private Sub ValidateEnvironment()

    If UCase$(ActiveSheet.Name) = UCase$("Master") Then
        ActivateUI
        MsgBox "Please select a sheet other than (Master)", vbOKOnly, MessageCaption
        End
    End If
    

End Sub

Public Function DisconnectRecordSet(rsRecordSetToDisconnect As ADODB.Recordset) As ADODB.Recordset
    Const FILE_NAME As String = "\agrData.xml"
    Dim sPersistedRecordsetFileName As String
    
    sPersistedRecordsetFileName = ActiveWorkbook.Path & FILE_NAME
    If LenB(Dir(sPersistedRecordsetFileName)) <> 0 Then
        Kill sPersistedRecordsetFileName
    End If
    
    
    rsRecordSetToDisconnect.Save sPersistedRecordsetFileName, adPersistXML
    rsRecordSetToDisconnect.Close
    
    Set rsRecordSetToDisconnect = Nothing
    Set rsRecordSetToDisconnect = New ADODB.Recordset

    rsRecordSetToDisconnect.CursorLocation = adUseClient
    ' rsRecordSetToDisconnect.LockType = adLockBatchOptimistic
    rsRecordSetToDisconnect.LockType = adLockOptimistic
    rsRecordSetToDisconnect.Open sPersistedRecordsetFileName
    Set DisconnectRecordSet = rsRecordSetToDisconnect
    
    
End Function


Public Sub mymacro1()
    MsgBox "Macro1 from a right click menu"
End Sub


Public Sub ThrowError()

    Dim ierr As Integer
    ierr = 1 / 0

End Sub


Public Function MessageCaption() As String
    MessageCaption = "(AGR) (" & Application.UserName & ") (" & Environ$("computername") & ")"
End Function
