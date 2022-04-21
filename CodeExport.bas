Option Explicit

Public Sub ExportCode()


    Dim vbcomp As VBComponent
    Dim fileWriter As New FileReadWrite
    Dim basePath As String
    Dim fullPath As String
    Dim code As String
    Dim ext As String
    
    basePath = ActiveWorkbook.Path & "\"
    
    If Left$(basePath, 3) = "C:\" Then
        basePath = "C:\Delete\VBA_Modes\Local\"
    Else
        basePath = "C:\Delete\VBA_Modes\Network\"
    End If
    
    
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents

        ' Debug.Print vbcomp.Name & ", " & Str(vbcomp.Type)
        
        Select Case vbcomp.Type
            
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
            
                If InStr(UCase$(vbcomp.Name), "MODULE") = 0 Then
                    fullPath = basePath & vbcomp.Name & GetExtension(vbcomp.Type)
                    code = vbcomp.CodeModule.Lines(1, vbcomp.CodeModule.CountOfLines)
                    code = CleanCrLfs(code)
                    fileWriter.WriteToFile fullPath, code
                End If
                
            Case Else
                ' do nothing
            
        End Select
        
    Next vbcomp


End Sub

Private Function GetExtension(fileType As Integer) As String

    Select Case fileType
    
        Case vbext_ct_StdModule
            GetExtension = ".bas"
                
        Case vbext_ct_ClassModule
            GetExtension = ".cls"
            
        Case vbext_ct_MSForm
            GetExtension = ".frm"
    
    End Select
    

End Function


Private Function CleanCrLfs(code As String) As String
    
    Do Until Right(code, Len(vbNewLine)) <> vbNewLine
        code = Left(code, (Len(code) - Len(vbNewLine)))
    Loop
    
    CleanCrLfs = code

End Function
