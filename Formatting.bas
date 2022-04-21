Option Explicit

Public Sub FormatCellOffsetColor(iRow As Integer, iCol As Integer)

    Range(Cells(iRow, iCol), Cells(iRow, iCol)).Font.Color = 255
    Range(Cells(iRow, iCol), Cells(iRow, iCol)).Interior.Color = 49407

End Sub


Public Sub FormatCellNone(iRow As Integer, iCol As Integer)

    Range(Cells(iRow, iCol), Cells(iRow, iCol)).Font.Color = 0
    Range(Cells(iRow, iCol), Cells(iRow, iCol)).Interior.ColorIndex = 2

End Sub

'' https://stackoverflow.com/questions/58282085/change-cell-color-when-it-selected-and-back-original-color-after-leaving-it
'' https://www.techrepublic.com/article/how-to-use-vba-to-change-the-active-cell-color-in-an-excel-sheet/
'Public Sub FormatBlock(ByVal Target As Excel.Range)
'
'    Static rngOld As Range
'    Static colorOld As Integer
'
'    On Error Resume Next
''
''    colorOld = Target.Interior.ColorIndex
''
''    Target.Interior.ColorIndex = 8 ' Cyan
''
''    rngOld.Interior.ColorIndex = colorOld
''
''    Set rngOld = Target
'
'End Sub
