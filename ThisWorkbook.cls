VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    
    Dim sheetName As String
    
    On Error GoTo Workbook_OpenErrHandler
    
    InitializeEnvironment
    End
'    ActivateUI
'    Worksheets("Master").Protect DrawingObjects:=True, contents:=True, Scenarios:=True
'    ClearLogFiles
    
    Application.Interactive = False

    frmStatus.Show vbModeless
    
    AddContextMenus ' this Sub exits early because this sub is not completed
    
    sheetName = GetSheetNameHighestWeekNum()
    
    frmStatus.UpdateMain "initializing Auto Run"
    
    GetDataAllDays sheetName
    
    Range("C3:C3").Select
    ActiveWorkbook.Save
    
    CheckMissingDataMsg
    
Workbook_OpenExitPoint:
    On Error Resume Next
    ActivateUI
    Unload frmStatus
    End
    
Workbook_OpenErrHandler:
    
    
    MsgBox "Error occurred in Workbook_Open" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, MessageCaption
    Resume Workbook_OpenExitPoint
    
    
    
End Sub


Private Function GetSheetNameHighestWeekNum() As String
    
    Dim num As Integer
    Dim highNum As Integer
    Dim highSheet As String
    Dim count As Integer
    
    
    count = 1
    
    Do Until count > ActiveWorkbook.Worksheets.count
    
        If UCase$(ActiveWorkbook.Worksheets(count).Name) <> UCase$("Master") Then
        
            num = GetNumChars(ActiveWorkbook.Worksheets(count).Name)
            
            If num > highNum Then
                highNum = num
                highSheet = ActiveWorkbook.Worksheets(count).Name
            End If
            
        End If
        count = count + 1
        
    Loop

    GetSheetNameHighestWeekNum = highSheet

End Function


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
   
    If InStr(Sh.Name, "(") And InStr(Sh.Name, ")") Then ' Runs when sheet is copied
        Dim weekNum As Integer
    
        Sh.Unprotect
        
        weekNum = GetNextSheetWeekNum
        Sh.Name = "Week " & Trim$(Str$(weekNum))
        
        Cells(1, 2).Value = Trim$(Str$(weekNum)) ' write the week number to the header cell

        
        End
        
    End If

    If UCase$(Sh.Name) = UCase$("Master") Then
        Exit Sub
    End If
    
    ValidWeekNumber Sh.Name
    
    ' Debug.Print Sh.Name
    
    If InStr(UCase$(Sh.Name), UCase$("Week")) <> 1 Then
        MsgBox "Sheet names should begin with the word (Week)", vbExclamation, "(" & Sh.Name & ") name is invalid", MessageCaption
    End If
    
End Sub


Private Sub ValidWeekNumber(val As String)
    Dim weekNum As Integer
  
    
    weekNum = GetNumChars(val)
    If weekNum > 0 Then
        If Worksheets(val).Range("B1").Value <> weekNum Then
            MsgBox "Week number = " & Worksheets(val).Range("B1").Value & vbCrLf & "Sheet Name = (" & val & ")", vbInformation, MessageCaption
            End
        End If
    End If

    
    
End Sub


Private Function GetNumChars(val As String) As Integer
    Dim retVal As String
    Dim char As String
    Dim count As Integer
    
    
    count = 1
    Do Until count > Len(val)
        char = Mid(val, count, 1)
        
        If IsNumeric(char) Then
            retVal = retVal & char
        End If
        
        count = count + 1
    Loop
    
    If IsNumeric(retVal) And retVal <> "0" Then
        GetNumChars = Int(retVal)
    Else
        If UCase$(val) <> UCase$("Master") Then
            MsgBox "Sheet (" & val & ")'s name is invalid" & vbCrLf & vbCrLf & "It appears to be missing the week number" & vbCrLf & vbCrLf & "Correct the name formatted like (Week ##)", vbOKOnly + vbExclamation, "Invalid sheet name (" & val & ")"
        End If
        
        GetNumChars = 0
    End If
    
    
    
    
End Function


Private Function MenuItemExists(itemName As String) As Boolean
    Dim itmCount As Integer
    
    MenuItemExists = False
    
    Do Until itmCount = Application.ShortcutMenus(xlWorksheetCell).MenuItems.count
        If UCase$(Application.ShortcutMenus(xlWorksheetCell).MenuItems.Item(1).Caption) = UCase$(itemName) Then
            MenuItemExists = True
            Exit Function
        End If
        itmCount = itmCount + 1
    Loop
    
    MenuItemExists = False

End Function

'
'    If UCase$(Sh.Name) = UCase$("Master") Then
'        Exit Sub
'    End If
'
'    If GetNumChars(Sh.Name) <= 0 Then
'        MsgBox "invalid name"
'    End If
'
'
'End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Target.count > 1 Then
'        Debug.Print Target(1).Value & ", (" & ActiveSheet.Name & ") (Col = " & Target.Column & ") (Row = " & Target.row & ")"
'        frmStatus.Label1.Caption = Target(1).Value & ", Col = " & Target.Column
        
    Else
'        Debug.Print Target.Value & ", (" & ActiveSheet.Name & ") (Col = " & Target.Column & ") (Row = " & Target.row & ")"
'        frmStatus.Label1.Caption = Target.Value & ", Col = " & Target.Column
    End If
End Sub

Private Function GetNextSheetWeekNum() As Integer
    Dim count As Integer
    Dim highNum As Integer
    
    count = 1
    Do Until count > ActiveWorkbook.Worksheets.count
        ' Debug.Print ActiveWorkbook.Worksheets(count).Name
        
        If UCase$(ActiveWorkbook.Worksheets(count).Name) <> UCase$("Master") Then
        
            If GetNumChars(ActiveWorkbook.Worksheets(count).Name) > highNum Then
                highNum = GetNumChars(ActiveWorkbook.Worksheets(count).Name)
            End If
            
        End If
        
        count = count + 1
    Loop
    
    GetNextSheetWeekNum = highNum + 1
    
End Function

Private Sub AddContextMenus()
    Dim MyMenu As Object
    
    
'    MsgBox "Skipping AddContextMenus"
    Exit Sub
    
    If MenuItemExists("Secondary AGR") = True Then
        Exit Sub
    End If
    
    Set MyMenu = Application.ShortcutMenus(xlWorksheetCell).MenuItems.AddMenu("Secondary AGR", 1)
    
    With MyMenu.MenuItems

        
        .Add "MyMacroZ", "MyMacro1", , 1, , ""

    End With
    
    Set MyMenu = Nothing

End Sub
