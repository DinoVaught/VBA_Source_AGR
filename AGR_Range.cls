Option Explicit

Private iShift3Row As Integer
Private iShift1Row As Integer
Private iShift2Row As Integer
Private iST_1_Col As Integer
Private iST_2_Col As Integer
Private iST_3_Col As Integer
Private iST_4_Col As Integer
Private iST_5_Col As Integer
Private iST_6_Col As Integer
Private iTotal_Col As Integer
Private iAGR_Col As Integer
Private iNET_Col As Integer

Private errMsg As String



Public Property Get ErrorMsg() As String
    ErrorMsg = errMsg
End Property

Public Property Get Shift3Row() As Integer
    Shift3Row = iShift3Row
End Property

Public Property Let Shift3Row(ByVal iValue As Integer)
    iShift3Row = iValue
End Property

Public Property Get Shift1Row() As Integer
    Shift1Row = iShift1Row
End Property

Public Property Let Shift1Row(ByVal iValue As Integer)
    iShift1Row = iValue
End Property

Public Property Get Shift2Row() As Integer
    Shift2Row = iShift2Row
End Property

Public Property Let Shift2Row(ByVal iValue As Integer)
    iShift2Row = iValue
End Property


Public Property Get ST_1_Col() As Integer
    ST_1_Col = iST_1_Col
End Property

Public Property Let ST_1_Col(ByVal iValue As Integer)
    iST_1_Col = iValue
End Property
Public Property Get ST_2_Col() As Integer
    ST_2_Col = iST_2_Col
End Property

Public Property Let ST_2_Col(ByVal iValue As Integer)
    iST_2_Col = iValue
End Property

Public Property Get ST_3_Col() As Integer
    ST_3_Col = iST_3_Col
End Property

Public Property Let ST_3_Col(ByVal iValue As Integer)
    iST_3_Col = iValue
End Property

Public Property Get ST_4_Col() As Integer
    ST_4_Col = iST_4_Col
End Property

Public Property Let ST_4_Col(ByVal iValue As Integer)
    iST_4_Col = iValue
End Property

Public Property Get ST_5_Col() As Integer
    ST_5_Col = iST_5_Col
End Property

Public Property Let ST_5_Col(ByVal iValue As Integer)
    iST_5_Col = iValue
End Property

Public Property Get ST_6_Col() As Integer
    ST_6_Col = iST_6_Col
End Property

Public Property Let ST_6_Col(ByVal iValue As Integer)
    iST_6_Col = iValue
End Property

Public Property Get Total_Col() As Integer
    Total_Col = iTotal_Col
End Property

Public Property Let Total_Col(ByVal iValue As Integer)
    iTotal_Col = iValue
End Property

Public Property Get AGR_Col() As Integer
    AGR_Col = iAGR_Col
End Property

Public Property Let AGR_Col(ByVal iValue As Integer)
    iAGR_Col = iValue
End Property

Public Property Get NET_Col() As Integer
    NET_Col = iNET_Col
End Property

Public Property Let NET_Col(ByVal iValue As Integer)
    iNET_Col = iValue
End Property

Public Sub Initialize(gageID As String, partNum As String, StartCol As Integer)
    Dim startRow As Integer
    
    errMsg = vbNullString
    
    startRow = 1
    
    If GageAndPartAreInSheet(gageID, partNum, startRow) = False Then
        LogMissingData ("Gage = " & gageID & vbCrLf & "Part Num = " & partNum & vbCrLf & "========================")

        Exit Sub
    End If
    
    iShift3Row = startRow
    iShift1Row = iShift3Row + 1
    iShift2Row = iShift3Row + 2
    
    
    iST_1_Col = StartCol
    iST_2_Col = iST_1_Col + 1
    iST_3_Col = iST_1_Col + 2
    iST_4_Col = iST_1_Col + 3
    iST_5_Col = iST_1_Col + 4
    iST_6_Col = iST_1_Col + 5
    iTotal_Col = iST_1_Col + 6
    iAGR_Col = iST_1_Col + 7
    iNET_Col = iST_1_Col + 8
    
End Sub


Private Function GageAndPartAreInSheet(gageID As String, partNum As String, ByRef rowNum As Integer) As Boolean

    Dim GageCell As Range
    Dim rowCount As Integer
    Dim count As Integer
    Dim testVal As String
    
    GageAndPartAreInSheet = False
    
    rowCount = ActiveSheet.UsedRange.Rows.count
    
    
'    If gageID = "G08" Then
'        gageID = gageID
'    End If

   
    gageID = UCase$(Trim$(gageID))
    partNum = UCase$(Trim$(partNum))
    
    
    Set GageCell = Sheets(ActiveSheet.Index).Range("B:B").Find(gageID)
    If GageCell Is Nothing Then
        GageAndPartAreInSheet = False
        Exit Function
    End If
    count = GageCell.Row

    Do Until count > rowCount

        testVal = UCase$(Trim$(Cells(count, 2).Value))
       
        If IsGageID(testVal) = True Then
        
            If testVal = gageID Then
                If Cells(count + 2, 2).Value = partNum Then
                    rowNum = count
                    GageAndPartAreInSheet = True
                    Exit Function
                End If
            End If
        End If
        
        count = count + 1
    Loop

    GageAndPartAreInSheet = False

End Function


Private Function IsGageID(val As String) As Boolean

    Const EXIT_RETURN_FALSE As Long = vbObjectError + 1000
    Const LETTER_G As String = "G"
    Const ALL_GAGE_IDS As String = "G01:G02:G03:G04:G05:G06:G07:G08:G09:G10:G10:G11:G12:G12:G13:G14:G15:G16:G17:G18:G18:G19:G19:G20:G21:G22:G23:G24:G25:G26:G27"

    On Error GoTo IsGageIDErrHandler
    
    val = Trim$(val)
    
    If val = vbNullString Then
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    val = UCase$(val)
    
    If InStr(1, ALL_GAGE_IDS, val) = 0 Then
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    If Len(val) <> 3 Then
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    If IsNumeric(val) = True Then ' Gage IDs are not Numeric
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    If Left$(val, 1) <> LETTER_G Then
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    If IsNumeric(Replace$(val, LETTER_G, vbNullString)) = False Then
        Err.Raise EXIT_RETURN_FALSE
    End If
    
    IsGageID = True
    
IsGageIDExitPoint:
    On Error Resume Next
    Exit Function
    
IsGageIDErrHandler:

    Select Case Err.Number
    
        Case EXIT_RETURN_FALSE
            IsGageID = False
    
           
        Case Else
            MsgBox "Error occurred in Formatting.IsGageID" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbInformation, MessageCaption
        
    End Select
    
    IsGageID = False
    Resume IsGageIDExitPoint

End Function
