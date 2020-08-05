Attribute VB_Name = "ExcelUtil"
Option Explicit

Private evtSht As New EvtSheet
Private evtBk As New EvtBook

'''エクセルを開く
'''読み取り専用指定が可能
Function TryOpenBook(str, rd) As Workbook
    On Error Resume Next
    Dim bk As Workbook
    For Each bk In Workbooks
        If bk.FullName = Trim(str) Then
            bk.Activate
            If rd Then
                bk.ChangeFileAccess xlReadOnly
            Else
                bk.ChangeFileAccess xlReadWrite
            End If
            
            Set TryOpenBook = bk
            Exit Function
        End If
    Next
    
    Set TryOpenBook = Workbooks.Open(str, readonly:=rd)
    TryOpenBook.Activate
    On Error GoTo 0
End Function

'''ショートカットキー設定
Sub SetShotCutKeys()
    Dim Row As Long
    Dim rdSht As Worksheet
    Set rdSht = ThisWorkbook.Sheets("ショートカット")
    For Row = 2 To 100
        If rdSht.Cells(Row, 1).Value = Empty Then
            Exit For
        End If
        
        Application.OnKey TranslateKey(rdSht.Cells(Row, 2).Text), rdSht.Cells(Row, 3).Text
    Next
End Sub

'''組み合わせキーを翻訳
'ex:"SHIFT+CTRL+ALT+R"→"+^%R"
Function TranslateKey(str) As String
    Dim strs
    strs = Split(str, "+")
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To UBound(strs)
        strTemp = strs(i)
        strTemp = Trim(strTemp)
        strTemp = Replace(strTemp, "SHIFT", "+")
        strTemp = Replace(strTemp, "CTRL", "^")
        strTemp = Replace(strTemp, "ALT", "%")
        'strTemp = Replace(strTemp, "ENTER", "~")
        TranslateKey = TranslateKey + strTemp
        
    Next
End Function


'''シート名検索
Sub SearchSheet()
    Dim resultArr As New Collection
    Dim bk As Workbook
    Set bk = ActiveWorkbook
    Dim searchStr As String
    
    searchStr = InputBox("シート名キーワードを入力してください（*で省略できます）：", "シート名入力", Constant.SheetSearchStr.Text)
    If searchStr <> Empty Then
        SheetSearchStr.Value = searchStr
        Dim sht As Worksheet
        For Each sht In bk.Sheets
            If StringUtil.RegExMatch("*" + searchStr + "*", sht.Name) Then
                resultArr.Add sht
            End If
            
        Next
        
        If resultArr.Count = 0 Then
            MsgBox "対象シートがみつかりませんでした。"
        ElseIf resultArr.Count = 1 Then
            resultArr(1).Activate
        Else
            SheetSelectionForm.setResultArr resultArr
            SheetSelectionForm.Show
        End If
    End If

End Sub

'''ブックパスコピー
Sub CopyBookPath()
    SysUtil.PutInClipboard ActiveWorkbook.FullName
    MsgBox "下記パスをコピーしました：" + vbCrLf + vbCrLf + ActiveWorkbook.FullName
End Sub

'''複数セルインプット
Sub MultiInput()

    Dim rng As Range
    Set rng = Selection
    On Error Resume Next
    Dim formula
    formula = rng(1, 1).Validation.Formula1
    On Error GoTo 0
    If formula <> Empty Then

        If Left(formula, 1) = "=" Then
            formula = Evaluate(formula)
            InputForm.ShowInputFormList rng, formula
        Else
            InputForm.ShowInputFormArr rng, Split(formula, ",")
        End If
        
    Else
        Dim inputStr As String
        
        inputStr = InputBox("入力内容を入力してください：", "複数セル一括入力", Constant.MultiInputString.Text)
        If inputStr <> Empty Then
            Constant.MultiInputString.Value = inputStr
            
            Dim cel As Range
            For Each cel In rng
                If cel.Rows.Hidden = False And cel.Columns.Hidden = False And (cel.Value = Empty Or Constant.MultiInputOverwrite.Value = "〇") Then
                    cel.Value = inputStr
                End If
            Next
        End If
    End If
End Sub

'''選択内容を合併したように見せる
Sub DoSelectionLikeCombined()
'
' Macro1 Macro
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'    With Selection.Font
'        .ThemeColor = xlThemeColorDark1
'    End With
End Sub

'''シート複数コピー
Sub MultiCopySheet()

    Dim num
    num = CInt(InputBox("何個分コピーするのか？"))
    
    Dim copyFromSht As Worksheet
    Set copyFromSht = ActiveSheet
    Dim newSht As Worksheet
    
    Dim i
    For i = 1 To num
        Call copyFromSht.Copy(After:=copyFromSht)
        Set newSht = ActiveSheet
        newSht.Name = CStr(CInt(copyFromSht.Name) + 1)
        Set copyFromSht = newSht
    Next

End Sub


'''正規表現でクリア
Sub ClearByRegExp()
    Dim cel As Range
    
    For Each cel In Selection
        If Not StringUtil.RegExMatch(Constant.ClearByRegRepSetting.Text, cel.Text) Then
            cel.Clear
        End If
    Next
End Sub

Function IsSelectedPCL(Row As Long, col As Integer) As Boolean
    IsSelectedPCL = False
    If Row = ActiveCell.Row Or col = ActiveCell.Column Then
        IsSelectedPCL = True
    ElseIf Trim(ActiveSheet.Cells(Row, ActiveCell.Column).Text) = "○" Then
        IsSelectedPCL = True
    End If
End Function


'test
Sub test()
    Dim vbc
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Debug.Print vbc.Name
    Next
End Sub

'''F1キーを無効にしたメッセージを表示
Sub DisableF1()
    Application.StatusBar = "F1キーを無効にしています。有効にしたい場合は[" & ThisWorkbook.Name & "]を開かずにエクセルを再起動してください。"
End Sub

'''リサーチパネルを非表示に
Sub DiableResearchPanel()
    Application.CommandBars("Research").Enabled = False
End Sub

'''選択セルのフォーマットを「HH:mm:ss.000」に
Sub SetHMS000()
    Selection.NumberFormatLocal = "hh:mm:ss.000"
End Sub

Sub printSheetNames()
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Sheets
        Debug.Print sht.Name & vbTab & sht.Visible
    Next
End Sub
