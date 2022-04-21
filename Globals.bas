Option Explicit
Private Const DATA_NOT_FOUND_FILE As String = "\agr_data_not_found.txt"
Public GageNotFoundFile As String
Private EnvironmentInitialized As Boolean
Public DayRange As DayRanges
Public FindGage As GageFinder
Public CellFormatter As SelectedCellColors

Public Sub InitializeEnvironment()

    If EnvironmentInitialized = True Then
        Exit Sub
    End If
    
    GageNotFoundFile = Application.ActiveWorkbook.Path & DATA_NOT_FOUND_FILE
    ActivateUI
    ClearLogFiles
    Worksheets("Master").Protect DrawingObjects:=True, contents:=True, Scenarios:=True
    
    EnvironmentInitialized = True
    
End Sub

Public Sub ActivateUI()
    On Error Resume Next
    Unload frmStatus
    Application.Interactive = True
    Application.Cursor = xlDefault
End Sub

Public Sub ClearLogFiles()

    On Error Resume Next
    Kill GageNotFoundFile

End Sub

Public Sub LogMissingData(data As String)
    
    Dim fileIO As New FileReadWrite
    Dim contents As String
    
    contents = fileIO.ReadWholeFile(GageNotFoundFile)
    
    If contents = vbNullString Then
        Dim header As String ' writing first record, add header data
        header = Now & vbCrLf
        header = header & "Sheet = " & Application.ActiveSheet.Name & vbCrLf
        header = header & "Could not find a match for the specified items" & vbCrLf
        
        data = header & vbCrLf & data

        fileIO.AppendToFile GageNotFoundFile, data
        Exit Sub
    End If
    
    If InStr(1, contents, data) = 0 Then
        fileIO.AppendToFile GageNotFoundFile, data
        Exit Sub
    End If
    
End Sub

Public Sub CheckMissingDataMsg()

    Dim fileIO As New FileReadWrite
    Dim contents As String
    
    contents = fileIO.ReadWholeFile(GageNotFoundFile)
    
    If Len(contents) > 0 Then
        MsgBox contents, vbInformation, MessageCaption
    End If

End Sub


Public Sub CopyToClip(data As String)
    On Error Resume Next
    
    'Dim varText As Variant
    Dim objCP As Object
    'varText = "Some copied text"
    Set objCP = CreateObject("HtmlFile")
    objCP.ParentWindow.ClipboardData.SetData "text", data
    
End Sub
