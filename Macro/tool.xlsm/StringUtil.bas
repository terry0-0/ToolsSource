Attribute VB_Name = "StringUtil"
Option Explicit

'''対象文字列をパターンで確認
'''「*」
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

'''対象文字列をパターンで確認
'''「*」
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

'''辞書置換
Sub DicReplace()
    
    If Not MsgBox("辞書置換するには時間がかかります、以下のを確認済みですか？" + vbCrLf + vbCrLf + "辞書設定が正しいこと" + vbCrLf + "対象文字列をコピーしたこと", vbDefaultButton2 + vbYesNo) = vbYes Then
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
            
            If dicSht.Name = "表紙" Or _
                    dicSht.Name = "更新履歴" Then
                    'dicSht.Name = "テーブル一覧" Then
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
    
    MsgBox "辞書置換完了しました、テキストエディタに貼り付けてください。"
End Sub
