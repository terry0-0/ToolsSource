Attribute VB_Name = "PCLTool"
Option Explicit

'''PCL系の動作モジュール
'''例：選択PCLの行列ハイライト表示

'''イベント付きブック
Private evtBk As New EvtBook
'''イベント付きシート
Private evtSht As New EvtSheet



'''SetSelectionGriding
'''チェックリストマトリクス確認用、対象条件をハイライト表示する
Sub SetSelectionGriding()
    Dim rngstr
    rngstr = Trim(SelectionGriding.Text)
     
    If rngstr = Empty Then
        rngstr = InputBox("Input target Range")
    End If
    If rngstr = Empty Then
        Exit Sub
    End If
     
    SelectionGriding.Value = rngstr
     

    Call UnSetSelectionGriding(ActiveSheet)
    Dim formulaStr
    Dim pclMarkStr
    formulaStr = "=OR(COLUMN()=CELL(""col"")"
    For Each pclMarkStr In Split(Constant.pclMark.Text, ",")
        formulaStr = formulaStr + ", INDIRECT(""R"" & ROW() & ""C"" & CELL(""col""), FALSE)=""" & pclMarkStr & """"
    Next
    formulaStr = formulaStr + ")"
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=OR(ROW()=CELL(""row""),COLUMN()=CELL(""col""), INDIRECT(""R"" & ROW() & ""C"" & CELL(""col""), FALSE)=""○"")", Formula2:=""
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=OR(COLUMN()=CELL(""col""), INDIRECT(""R"" & ROW() & ""C"" & CELL(""col""), FALSE)=""○"")", Formula2:=""
    ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:=formulaStr, Formula2:=""
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=tool.xlsm!IsSelectedPCL(ROW(),COLUMN())", Formula2:=""
    ActiveSheet.Range(rngstr).FormatConditions(1).SetFirstPriority
    ActiveSheet.Range(rngstr).FormatConditions(1).Interior.Color = 5296274
    ActiveSheet.Range(rngstr).FormatConditions(1).StopIfTrue = False

 
    Set evtSht.evtSht = ActiveSheet
    Set evtBk.evtBk = ActiveWorkbook
End Sub

'''SetPCLGriding
Sub SetPCLGriding()
    Dim rngstr
    rngstr = Trim(SelectionGriding.Text)
     
    If rngstr = Empty Then
        rngstr = InputBox("Input target Range")
    End If
    If rngstr = Empty Then
        Exit Sub
    End If
     
    SelectionGriding.Value = rngstr
     

    Call UnSetSelectionGriding(ActiveSheet)
    Dim formulaStr
    Dim pclMarkStr
    formulaStr = "=OR(COLUMN()=" & CStr(Selection.Column)
    For Each pclMarkStr In Split(Constant.pclMark.Text, ",")
        formulaStr = formulaStr + ", INDIRECT(""R"" & ROW() & ""C" & CStr(Selection.Column) & """, FALSE)=""" & pclMarkStr & """"
    Next
    formulaStr = formulaStr + ")"
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=OR(ROW()=CELL(""row""),COLUMN()=CELL(""col""), INDIRECT(""R"" & ROW() & ""C"" & CELL(""col""), FALSE)=""○"")", Formula2:=""
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=OR(COLUMN()=CELL(""col""), INDIRECT(""R"" & ROW() & ""C"" & CELL(""col""), FALSE)=""○"")", Formula2:=""
    ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:=formulaStr, Formula2:=""
    'ActiveSheet.Range(rngstr).FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=tool.xlsm!IsSelectedPCL(ROW(),COLUMN())", Formula2:=""
    ActiveSheet.Range(rngstr).FormatConditions(1).SetFirstPriority
    ActiveSheet.Range(rngstr).FormatConditions(1).Interior.Color = 65535
    ActiveSheet.Range(rngstr).FormatConditions(1).StopIfTrue = False

 
    Set evtSht.evtSht = Nothing
    Set evtBk.evtBk = ActiveWorkbook
End Sub


Sub UnSetSelectionGriding(sht As Worksheet)
    Dim cnd As FormatCondition
'    For Each cnd In sht.Cells.FormatConditions
'        If cnd.Formula1 = "=OR(ROW()=CELL(" & """" & "row" & """" & "),COLUMN()=CELL(" & """" & "col" & """" & "))" Then
            sht.Cells.FormatConditions.Delete
'            Exit For
'        End If
'    Next
    Set evtSht.evtSht = Nothing
    Set evtBk.evtBk = Nothing
End Sub
