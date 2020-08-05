Attribute VB_Name = "AutoGen"
Option Explicit

'''自動生成用コール
Sub AutoGen()
    Dim outputStr As String
    Dim outputStrTemp As String
    Dim rdSht As Worksheet, tpltSht As Worksheet
    Set rdSht = ThisWorkbook.Sheets("自動生成")
    Set tpltSht = ThisWorkbook.Sheets("自動生成テンプレート")
    Dim Row As Long, tempRow As Long
    For Row = 2 To rdSht.Rows.Count
        If rdSht.Cells(Row, 1).Text = Empty Then
            Exit For
        End If
        
        tempRow = FindTemplateRow(tpltSht, rdSht.Cells(Row, 1).Text)
        If tempRow > 0 Then
            outputStrTemp = tpltSht.Cells(tempRow, 2).Text
            
            Dim keyCnt As Integer
            For keyCnt = 1 To tpltSht.Cells(tempRow, 4).Value
                outputStrTemp = Replace(outputStrTemp, tpltSht.Cells(tempRow, 3).Text & Format(keyCnt, "00"), rdSht.Cells(Row, 1 + keyCnt).Text)
            Next
            outputStrTemp = Replace(outputStrTemp, vbLf, vbCrLf)
            outputStr = outputStr & outputStrTemp
        End If
    Next
    'Debug.Print outputStr
    SysUtil.PutInClipboard outputStr
    MsgBox "自動生成完了！"
End Sub

'''自動生成用テンプレートを呼び出す
Function FindTemplateRow(tpltSht As Worksheet, id As String)
    FindTemplateRow = 0
    Dim Row As Long
    For Row = 2 To tpltSht.Rows.Count
        If tpltSht.Cells(Row, 1).Text = Empty Then
            Exit For
        End If
        
        If id = tpltSht.Cells(Row, 1).Text Then
            FindTemplateRow = Row
            Exit Function
        End If
    Next
End Function
