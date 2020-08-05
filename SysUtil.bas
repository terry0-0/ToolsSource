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


'' WindowsAPI�錾
'' �eAPI�̏ڍׂ� (http://wisdom.sakura.ne.jp/system/winapi/win32/win90.html) ���Q�Ƃ̂���
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

' WindowsAPI�J�Ďg�p����萔�錾
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_UNICODETEXT As Long = &HD


Public Sub test()

    ' �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
    Call SetClipboard("�e�X�g�e�L�X�g")

    ' �N���b�v�{�[�h�̃e�L�X�g���擾���ăC�~�f�B�G�C�g�E�C���h�E�ɏo��
    Debug.Print GetClipboard

End Sub

' ��������N���b�v�{�[�h�ɃR�s�[���܂�
' �{������ DataObject ���g�p���邱�ƂňӐ}���Ȃ������񂪃N���b�v�{�[�h�ɓ\������Ƃ����邽��
' DataObject ���g�p������ WindowsAPI ���g�p���������Ŏ�������
' �܂��A�������̓������̃��b�N�����{���邽�߁A�{�������ŋ����I�����Ȃ��悤�ɂ��邱��
Public Sub PutInClipboard(ByRef sUniText As String)

    ' �ϐ��錾
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long

    ' �N���b�v�{�[�h���J�� (�J���Ȃ������ꍇ�̓G���[���X���[)
    If OpenClipboard(0&) = 0 Then
        ' ���͊J���Ă����ꍇ�ɔ����ĕ��鏈�����Ăяo��
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�N���b�v�{�[�h���J���܂���ł���"
    End If

    ' �N���b�v�{�[�h����ɂ��� (���s�����ꍇ�̓N���b�v�{�[�h����ăG���[���X���[)
    If EmptyClipboard = 0 Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�N���b�v�{�[�h����ɂł��܂���ł���"
    End If

    ' �m�ۂ��郁�����̈���擾 (�I�[�������C�h������ �����m�ۂ��邽�߂�2�o�C�g�����m��)
    iLen = LenB(sUniText) + 2&

    ' �q�[�v����w�肳�ꂽ�o�C�g���̃��������m�� (���s�����ꍇ�̓N���b�v�{�[�h����ăG���[���X���[)
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    If iStrPtr = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�������̊m�ۂɎ��s���܂���"
    End If

    ' �O���[�o���������I�u�W�F�N�g�����b�N���A�������u���b�N�̐擪�ւ̃|�C���^���擾
    iLock = GlobalLock(iStrPtr)
    If iLock = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�������̃��b�N�Ɏ��s���܂���"
    End If

    ' �w�肳�ꂽ��������A�������ɃR�s�[
    If lstrcpy(iLock, StrPtr(sUniText)) = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�������̃R�s�[�Ɏ��s���܂���"
    End If

    ' �O���[�o���������I�u�W�F�N�g�̃��b�N�J�E���g�����炵�܂�
    ' 0�ȊO���ԋp���ꂽ�ꍇ�̓��b�N���������Ȃ������Ƃ��ăG���[���X���[���܂�
    If GlobalUnlock(iStrPtr) <> 0 Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�������̃A�����b�N�Ɏ��s���܂���"
    End If

    ' �N���b�v�{�[�h�ɁA�w�肳�ꂽ�f�[�^�`���Ńf�[�^���i�[
    If SetClipboardData(CF_UNICODETEXT, iStrPtr) = Null Then
        On Error Resume Next
        Call CloseClipboard
        On Error GoTo 0
        Err.Raise 1000, , "�N���b�v�{�[�h�Ƀf�[�^���i�[�ł��܂���ł���"
    End If

    ' �N���b�v�{�[�h�����
    If CloseClipboard = 0 Then
        Err.Raise 1000, , "�N���b�v�{�[�h���N���[�Y�ł��܂���ł���"
    End If

End Sub

''''��������N���b�v�{�[�h�Ɋi�[
'Sub PutInClipboard(str)
'    'Dim obj As New DataObject
'    Dim obj
'    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    With obj
'        .SetText str
'        .PutInClipboard
'    End With
'End Sub

''''��������N���b�v�{�[�h����擾
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

'''��������N���b�v�{�[�h����擾
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
