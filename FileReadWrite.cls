Option Explicit

Public Sub AppendToFile(fileName As String, data As String)

    Dim fileID As Integer
    Const EXIT_SUB As Long = vbObjectError + 1000
    
    On Error GoTo AppendToFileErrHandler
    
    fileID = FreeFile
    ' Open fileName For Input As FreeFile
    
    If Dir$(fileName) = vbNullString Then
        Open fileName For Output As fileID
        Print #fileID, data
        Err.Raise EXIT_SUB
    End If
   
    Open fileName For Append As fileID
    Print #fileID, data
    

AppendToFileExitPoint:
    On Error Resume Next
    Close fileID
    Reset
    Exit Sub
    
AppendToFileErrHandler:

    Select Case Err.Number
    
            Case EXIT_SUB
                ' do nothing
            
            Case Else
                MsgBox "Error occurred in Globals.AppendToFile" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, MessageCaption
        
    End Select
    
    Resume AppendToFileExitPoint


End Sub

Public Sub WriteToFile(fileName As String, data As String)

    Dim fileID As Integer
    
    On Error GoTo WriteToFileErrHandler
    fileID = FreeFile

    Open fileName For Output As fileID
    Print #fileID, data
    
WriteToFileExitPoint:
    On Error Resume Next
    Close fileID
    Reset
    Exit Sub

WriteToFileErrHandler:

    Select Case Err.Number
    
           ' Case EXIT_SUB
                ' do nothing
            
            Case Else
                MsgBox "Error occurred in Globals.WriteToFile" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, MessageCaption
        
    End Select
    
    Resume WriteToFileExitPoint
End Sub


Public Function ReadWholeFile(fileName As String) As String


    Dim fileID As Integer
    Const FILE_NOT_FOUND As Long = vbObjectError + 1000
    
    On Error GoTo ReadWholeFileErrHandler
    
    ReadWholeFile = vbNullString
    
    If Dir$(fileName) = vbNullString Then
        Err.Raise FILE_NOT_FOUND
    End If
   
    fileID = FreeFile
    Open fileName For Input As FreeFile
    
    ReadWholeFile = Input(LOF(fileID), fileID)

ReadWholeFileExitPoint:
    On Error Resume Next
    Close fileID
    Reset
    Exit Function
    
ReadWholeFileErrHandler:

    Select Case Err.Number
    
            Case FILE_NOT_FOUND
                ReadWholeFile = vbNullString
            
            Case Else
                MsgBox "Error occurred in Globals.ReadWholeFile" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, MessageCaption
        
    End Select
    
    Resume ReadWholeFileExitPoint


End Function
