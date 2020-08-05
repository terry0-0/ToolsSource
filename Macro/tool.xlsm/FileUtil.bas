Attribute VB_Name = "FileUtil"
Option Explicit

'''ファイルを開く
Sub OpenFile()
    Dim rdRng As Range
    Set rdRng = Selection
    Dim rng As Range
    
    For Each rng In rdRng
        If rng.Value <> Empty Then
            If Strings.LCase(Left(Trim(rng.Text), 4)) = "http" Then
                If rng.Font.Italic = True Then
                    '斜体のURLはIEで開く
                    RunCommand "cmd /c,start " + rng.Text
                Else
                    '普通のURLはEDGEで開く
                    RunCommand "cmd /c,start microsoft-edge:" + rng.Text
                End If
            ElseIf (LCase(Right(Trim(rng.Text), 4)) = "xlsx" Or _
                LCase(Right(Trim(rng.Text), 4)) = "xlsm" Or _
                LCase(Right(Trim(rng.Text), 3)) = "xls") Then
                ExcelUtil.TryOpenBook rng.Text, rng.Font.Italic
            ElseIf (LCase(Right(Trim(rng.Text), 4)) = "docx" Or _
                LCase(Right(Trim(rng.Text), 3)) = "doc") _
                And rng.Font.Italic = True Then
                'RunCommand "winword.exe /f " + rng.Text
                WordUtil.TryOpenWord rng.Text, True
            Else
                RunCommand "cmd /c,start " + rng.Text
            End If
        End If
    Next
End Sub

'''ファイル・コマンドを実行
Sub RunCommand(str)
    Dim rc As Long
    rc = Shell(str, vbNormalFocus)
    If rc = 0 Then MsgBox "起動に失敗しました"
End Sub

'''命令を実行し、標準出力を取得
Function RunCmdAndGetOutput(sCmd) As String
    Dim wsh, wExec, Result As String
    Set wsh = CreateObject("WScript.Shell")         ''(1)
    
    Set wExec = wsh.Exec(sCmd)  ''(3
   
    Dim s As String
    Dim sLine As String
    While Not wExec.StdOut.AtEndOfStream
        sLine = wExec.StdOut.ReadLine
        If sLine <> "" Then
            RunCmdAndGetOutput = RunCmdAndGetOutput & sLine & vbCrLf
        End If
    Wend

    Set wExec = Nothing
    Set wsh = Nothing

End Function
