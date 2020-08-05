Attribute VB_Name = "StringUtil"
Option Explicit

'''�Ώە�������p�^�[���Ŋm�F
'''�u*�v
Function RegExMatch(pattern, str) As Boolean
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .pattern = Replace(pattern, "*", ".*")
        .IgnoreCase = True
        .Global = True
    End With
    RegExMatch = RE.test(str)
End Function

'''�Ώە�������p�^�[���Ŋm�F
'''�u*�v
Function RegExReplace(pattern, str, repStr)
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .pattern = Replace(pattern, "*", ".*")
        .IgnoreCase = True
        .Global = True
    End With
    RegExReplace = RE.Replace(str, repStr)
End Function

Sub testRegexWholeWord()
    Dim str
    str = "A cat is not a dog.pig"
    Debug.Print RegExMatch("\bpig\b", str)
End Sub

'''�����u��
Sub DicReplace()
    
    If Not MsgBox("�����u������ɂ͎��Ԃ�������܂��A�ȉ��̂��m�F�ς݂ł����H" + vbCrLf + vbCrLf + "�����ݒ肪����������" + vbCrLf + "�Ώە�������R�s�[��������", vbDefaultButton2 + vbYesNo) = vbYes Then
        Exit Sub
    End If
    
    Dim str As String
    'str = SysUtil.GetFromClipboard
    Dim rng
    For Each rng In Selection
        str = rng.Text
        
        Dim dicShtList
        Dim dicBk As Workbook
        Set dicBk = ExcelUtil.TryOpenBook(Constant.ReplaceDicBook.Text, True)
        If Trim(Constant.ReplaceDicSht.Text) <> Empty Then
            Set dicShtList = New Collection
            dicShtList.Add dicBk.Sheets(Constant.ReplaceDicSht.Text)
        Else
            Set dicShtList = dicBk.Sheets
        End If
        
        Dim dicSht As Worksheet
        Dim searchCols, repCols
        Dim emptyRows As Integer
        For Each dicSht In dicShtList
            searchCols = Split(Constant.ReplaceDicSearchCol, ",")
            repCols = Split(Constant.ReplaceDicRepCol, ",")
            emptyRows = 0
            
            If dicSht.Name = "�\��" Or _
                    dicSht.Name = "�X�V����" Then
                    'dicSht.Name = "�e�[�u���ꗗ" Then
                GoTo NEXT_0
            End If
            
            Dim i As Long, j As Long
            For i = Constant.ReplaceDicStartRow.Value To dicSht.Rows.Count
                If dicSht.Cells(i, 1).Text = Empty Then
                    emptyRows = emptyRows + 1
                    If emptyRows > 100 Then
                        Exit For
                    End If
                Else
                    emptyRows = 0
                    For j = 0 To UBound(searchCols)
                        str = RegExReplace("\b" & dicSht.Cells(i, CInt(searchCols(j))).Text & "\b", str, dicSht.Cells(i, CInt(repCols(j))).Text)
                    Next
                End If
                
        
            Next
NEXT_0:
        Next
        rng.Value = str
    Next
    SysUtil.PutInClipboard str
    
    MsgBox "�����u���������܂����A�e�L�X�g�G�f�B�^�ɓ\��t���Ă��������B"
End Sub
