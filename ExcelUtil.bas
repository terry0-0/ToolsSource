Attribute VB_Name = "ExcelUtil"
Option Explicit

Private evtSht As New EvtSheet

'''Try to open a workbook, if already opened,set it to active.
Function TryOpenWorkbook(path, rdOnly As Boolean) As Workbook
    Dim bk As Workbook
    For Each bk In Workbooks
        If bk.FullName = path Then
            bk.Activate
            Set TryOpenWorkbook = bk
            Exit Function
        End If
    Next
    
    Set TryOpenWorkbook = Workbooks.Open(path, rdOnly)
End Function

'''SetShortcutKeys
Sub SetShortcutKeys()
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Sheets("ShortcutKeys")
    Dim row
    For row = 2 To sht.Rows.Count
        If sht.Cells(row, 2).Text = Empty Then
            Exit For
        End If
        'Application.MacroOptions Macro:=sht.Cells(row, 3).Text, ShortcutKey:=""
        Debug.Print "OnKey:" + StringUtil.TransferShorcutKey(sht.Cells(row, 2).Text) + "," + sht.Cells(row, 3).Text
        Application.OnKey StringUtil.TransferShorcutKey(sht.Cells(row, 2).Text), sht.Cells(row, 3).Text
    Next
End Sub

'''MultiInput
Sub MultiInput()
Attribute MultiInput.VB_ProcData.VB_Invoke_Func = " \n"
    Dim str
    str = Constant.MultiInput.Text
    Dim rng As Range
    Dim targetRng As Range
    Set targetRng = Selection
    Set rng = targetRng.Cells(1, 1)
    Dim formula As String
    On Error Resume Next
    formula = rng.formula
    On Error GoTo 0
    If Left(formula, 1) = "=" Then
    ElseIf formula <> Empty Then
    End If
    
    
End Sub


'''CopySheetName
Sub CopySheetName()
Attribute CopySheetName.VB_ProcData.VB_Invoke_Func = " \n"
    SystemUtil.SetClipboard ActiveWorkbook.FullName
    MsgBox "Copy the workbook path:" + vbCrLf + vbCrLf + ActiveWorkbook.FullName
End Sub

'''FindSheet
Sub FindSheet()
    Dim resultList As New Collection
    Dim str
    str = Constant.FindSheet.Text
    str = InputBox("Sheet name to find:", Default:=str)
    If str <> "" Then
        Constant.FindSheet.Value = str
        Dim sht As Worksheet
        For Each sht In ActiveWorkbook.Sheets
            If RegTest(sht.Name, Replace(str, "*", ".*")) Then
                resultList.Add sht.Name
            End If
        Next
        
        If resultList.Count = 0 Then
            MsgBox "No result!"
        ElseIf resultList.Count = 1 Then
            ActiveWorkbook.Sheets(resultList(1)).Activate
        Else
            
        End If
    End If
    

End Sub

'''FindTargetRow
Function FindTargetRow(str As String, sht As Worksheet, col As Integer) As Long
    Dim row As Long
    For row = 1 To sht.Rows.Count
        If sht.Cells(row, col).Text = str Then
            FindTargetRow = row
            Exit Function
        End If
    Next
End Function

'''SetSelectionGriding
Sub SetSelectionGriding()
    Dim rngStr
    rngStr = Trim(SelectionGriding.Text)
    
    If rngStr = Empty Then
        rngStr = InputBox("Input target Range")
    End If
    If rngStr = Empty Then
        Exit Sub
    End If
    
    SelectionGriding.Value = rngStr
    

    ActiveSheet.Range(rngStr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=OR(ROW()=CELL(" & """" & "row" & """" & "),COLUMN()=CELL(" & """" & "col" & """" & "))", Formula2:=""
    ActiveSheet.Range(rngStr).FormatConditions(1).SetFirstPriority
    ActiveSheet.Range(rngStr).FormatConditions(1).Interior.Pattern = xlPatternSolid
    ActiveSheet.Range(rngStr).FormatConditions(1).Interior.PatternColorIndex = -4105
    ActiveSheet.Range(rngStr).FormatConditions(1).Interior.PatternTintAndShade = 0
    ActiveSheet.Range(rngStr).FormatConditions(1).Interior.Color = 65535
    ActiveSheet.Range(rngStr).FormatConditions(1).Interior.TintAndShade = 0
    ActiveSheet.Range(rngStr).FormatConditions(1).StopIfTrue = False

    Set evtSht.evtSht = ActiveSheet
End Sub


'// 指定ワークブックに指定フォルダ配下のモジュールをインポートする
'// 引数１：ワークブック
'// 引数２：モジュール格納フォルダパス
Sub ImportAll(a_TargetBook As Workbook, a_sModulePath As String)
    On Error Resume Next
    
    Dim oFso        As New FileSystemObject     '// FileSystemObjectオブジェクト
    Dim sArModule() As String                   '// モジュールファイル配列
    Dim sModule                                 '// モジュールファイル
    Dim sExt        As String                   '// 拡張子
    Dim iMsg                                    '// MsgBox関数戻り値
    
    iMsg = MsgBox("同名のモジュールは上書きします。よろしいですか？", vbOKCancel, "上書き確認")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// 全モジュールのファイルパスを取得
    Call searchAllFile(a_sModulePath, sArModule)
    
    '// 全モジュールをループ
    For Each sModule In sArModule
        '// 拡張子を小文字で取得
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// 拡張子がcls、frm、basのいずれかの場合
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// 同名モジュールを削除
            Call a_TargetBook.VBProject.VBComponents.Remove(a_TargetBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// モジュールを追加
            Call a_TargetBook.VBProject.VBComponents.Import(sModule)
            '// Import確認用ログ出力
            Debug.Print sModule
        End If
    Next
End Sub

'''ExportAll
Sub ExportAll()
    Dim module
    Dim extension
    Dim sPath
    Dim sFilePath
    
    sPath = ThisWorkbook.path
    
    For Each module In ThisWorkbook.VBProject.VBComponents
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ElseIf (module.Type = vbext_ct_MSForm) Then
            extension = "frm"
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        Else
            GoTo CONTINUE
        End If
        
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub
