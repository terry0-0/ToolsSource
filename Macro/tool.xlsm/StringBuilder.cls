VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buff() As String
Private lengh As Long

Public Sub Append(str)
    ReDim Preserve buff(lengh)
    buff(lengh) = str
    lengh = lengh + 1

End Sub

Public Function ToString() As String
    ToString = Join(buff, "")
End Function

Public Sub Clear()
    ReDim buff(0)
End Sub
