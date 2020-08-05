Attribute VB_Name = "SysUtil"
Option Explicit



'#If VBA7 And Win64 Then
'Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
'Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
'Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
'Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
'Private Declare PtrSafe Function MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr) As LongPtr
'Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
'#Else
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
'Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'#End If

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long



Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long

Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long



Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long


'' WindowsAPI宣言
'' 各APIの詳細は (http://wisdom.sakura.ne.jp/system/winapi/win32/win90.html) を参照のこと
'Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
'Private Declare Function CloseClipboard Lib "user32.dll" () As Long
'Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
'Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
'Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
'Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long

' WindowsAPI゛て使用する定数宣言
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_UNICODETEXT As Long = &HD


Public Sub test()

    ' クリップボードにテキストを設定
    Call SetClipboard("テストテキスト")

    ' クリップボードのテキストを取得してイミディエイトウインドウに出力
    Debug.Print GetClipboard

End Sub

' 文字列をクリップボードにコピーします
' 本処理は DataObject を使用することで意図しない文字列がクリップボードに貼りつくことがあるため
' DataObject を使用せずに WindowsAPI を使用した処理で実装する
' また、処理中はメモリのロックを実施するため、本処理中で強制終了しないようにすること
Public Sub PutInClipboard(ByRef sUniText As String)

    ' 変数宣言
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long

    ' クリップボードを開く (開けなかった場合はエラーをスロー)
    If OpenClipboard(0&) = 0 Then
        ' 実は開けていた場合に備えて閉じる処理を呼び出す
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "クリップボードを開けませんでした"
    End If

    ' クリップボードを空にする (失敗した場合はクリップボードを閉じてエラーをスロー)
    If EmptyClipboard = 0 Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "クリップボードを空にできませんでした"
    End If

    ' 確保するメモリ領域を取得 (終端文字ワイド文字列 分も確保するために2バイト多く確保)
    iLen = LenB(sUniText) + 2&

    ' ヒープから指定されたバイト数のメモリを確保 (失敗した場合はクリップボードを閉じてエラーをスロー)
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    If iStrPtr = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "メモリの確保に失敗しました"
    End If

    ' グローバルメモリオブジェクトをロックし、メモリブロックの先頭へのポインタを取得
    iLock = GlobalLock(iStrPtr)
    If iLock = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "メモリのロックに失敗しました"
    End If

    ' 指定された文字列を、メモリにコピー
    If lstrcpy(iLock, StrPtr(sUniText)) = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "メモリのコピーに失敗しました"
    End If

    ' グローバルメモリオブジェクトのロックカウントを減らします
    ' 0以外が返却された場合はロックが解放されなかったとしてエラーをスローします
    If GlobalUnlock(iStrPtr) <> 0 Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "メモリのアンロックに失敗しました"
    End If

    ' クリップボードに、指定されたデータ形式でデータを格納
    If SetClipboardData(CF_UNICODETEXT, iStrPtr) = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "クリップボードにデータを格納できませんでした"
    End If

    ' クリップボードを閉じる
    If CloseClipboard = 0 Then
        Err.Raise 1000, , "クリップボードをクローズできませんでした"
    End If

End Sub

''''文字列をクリップボードに格納
'Sub PutInClipboard(str)
'    'Dim obj As New DataObject
'    Dim obj
'    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    With obj
'        .SetText str
'        .PutInClipboard
'    End With
'End Sub

''''文字列をクリップボードから取得
'Function GetFromClipboard() As String
'
'    'Dim obj As New DataObject
'    Dim obj
'    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    With obj
'        .GetFromClipboard
'        GetFromClipboard = obj.GetText
'    End With
'End Function

'''文字列をクリップボードから取得
Public Function GetFromClipboard() As String
On Error GoTo ErrHandler
     
    Dim i As Long
    Dim Format As Long
    Dim Data() As Byte
#If VBA7 And Win64 Then
    Dim hMem As LongPtr
    Dim Size As LongPtr
    Dim p As LongPtr
#Else
    Dim hMem As Long
    Dim Size As Long
    Dim p As Long
#End If
     
    Call OpenClipboard(0)
    hMem = GetClipboardData(RegisterClipboardFormat("Text"))
    If hMem = 0 Then
        Call CloseClipboard
        Exit Function
    End If
     
    Size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(0 To CLng(Size) - CLng(1))
#If VBA7 And Win64 Then
    Call MoveMemory(Data(0), ByVal p, Size)
#Else
    Call MoveMemory(CLng(VarPtr(Data(0))), p, Size)
#End If
    Call GlobalUnlock(hMem)
     
    Call CloseClipboard
     
    For i = 0 To CLng(Size) - CLng(1)
        If Data(i) = 0 Then
            Data(i) = Asc(" ")
        End If
    Next i
     
    GetFromClipboard = StrConv(Data, vbUnicode)
    Exit Function
     
ErrHandler:
    Call CloseClipboard
    GetFromClipboard = ""
End Function
