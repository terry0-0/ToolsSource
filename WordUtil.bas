Attribute VB_Name = "WordUtil"
Option Explicit


Function TryOpenWord(path, readonly)
    Dim wrd
    Set wrd = Nothing
    On Error Resume Next
    Set wrd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wrd Is Nothing Then
        Set wrd = CreateObject("Word.Application")
    End If
    wrd.Visible = True
    Dim doc
    Set doc = wrd.Documents.Open(path, readonly:=readonly)
    doc.Activate
    Set wrd = Nothing
End Function

