Option Explicit

Private PartNums As Collection
Private pn As New partNum

Public Sub Initialize()

    Set PartNums = New Collection
    Dim prtCount As Integer
'    Dim pn As New partNum

        
    AddPartNumber "038"
    AddPartNumber "081"
    AddPartNumber "086"
    AddPartNumber "260"
    AddPartNumber "262"
    AddPartNumber "275"
    AddPartNumber "289"
    AddPartNumber "320"
    AddPartNumber "325"
    AddPartNumber "386"
    AddPartNumber "387"
    AddPartNumber "452"
    AddPartNumber "480"
    AddPartNumber "487"
    AddPartNumber "538"
    AddPartNumber "648"
    AddPartNumber "658"
    AddPartNumber "680"
    AddPartNumber "690"
    AddPartNumber "691"
    AddPartNumber "712"
    AddPartNumber "812"
    AddPartNumber "881"
    AddPartNumber "936"
    AddPartNumber "985"
    AddPartNumber "987"
    

    prtCount = 1
    Do While prtCount < PartNums.count + 1
        Debug.Print PartNums(prtCount).partNum
        prtCount = prtCount + 1
    Loop

'    For Each itm In PartNums.count
'        Debug.Print "itm."
'    Next



End Sub


Private Sub AddPartNumber(partNumber As String)

    Set pn = New partNum
    pn.partNum = partNumber
    PartNums.Add pn

End Sub


Public Sub GetPartNums()

    

End Sub
