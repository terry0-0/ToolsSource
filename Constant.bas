Attribute VB_Name = "Constant"
Option Explicit

'''MultiInput
Function MultiInput() As Range
    Set MultiInput = ThisWorkbook.Sheets("Setting").Cells(5, 2)
End Function

'''FindSheet
Function FindSheet() As Range
    Set FindSheet = ThisWorkbook.Sheets("setting").Cells(7, 2)
End Function

'''SelectionGriding
Function SelectionGriding() As Range
    Set SelectionGriding = ThisWorkbook.Sheets("setting").Cells(9, 2)
End Function

