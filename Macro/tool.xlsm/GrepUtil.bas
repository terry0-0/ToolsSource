Attribute VB_Name = "GrepUtil"
Option Explicit

Sub GrepFile()
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    Dim sel, rng As Range
    Set sel = Selection
    
    Dim cmd
    Dim settingSht As Worksheet
    Set settingSht = ThisWorkbook.Sheets("Grep�ݒ�")
    For Each rng In sel
        'sakura.exe
        cmd = settingSht.Cells(1, 2).Text
        cmd = cmd & " -GREPMODE "
        '��������
        cmd = cmd & "-GKEY=""" & Replace(rng.Text, ".java", "") & """"
        cmd = cmd & " "
        '�t�@�C���p�^�[��
        cmd = cmd & "-GFILE=""" & settingSht.Cells(2, 2).Text & """"
        cmd = cmd & " "
        '�t�H���_
        cmd = cmd & "-GFOLDER=""" & settingSht.Cells(3, 2).Text & """"
        cmd = cmd & " "
        'GREP�ݒ�
        cmd = cmd & "-GOPT:""" & settingSht.Cells(4, 2).Text & """"
        cmd = cmd & " "
        'Debug.Print cmd & vbCrLf
        
        Dim strResult
        strResult = FileUtil.RunCmdAndGetOutput(cmd)
        
        Dim str
        For Each str In Split(strResult, vbCrLf)
            'Debug.Print str + vbCrLf
            If Left(str, 1) = "��" And StringUtil.RegExMatch(settingSht.Cells(5, 2).Text, str) Then
                sht.Cells(rng.Row, settingSht.Cells(6, 2).Value).Value = Replace(StringUtil.RegExReplace("^.*\\", str, ""), """", "")
                Exit For
            End If
        Next
    Next
    
End Sub
