VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EvtBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''�C�x���g�t�����[�N�u�b�N
Public WithEvents evtBk As Workbook
Attribute evtBk.VB_VarHelpID = -1

'''�ۑ�����O�ɏ����������N���A
Private Sub evtBk_BeforeSave(ByVal SaveAsUI As Boolean, _
        Cancel As Boolean)
    Dim sht As Worksheet
    For Each sht In evtBk.Sheets
        UnSetSelectionGriding sht
    Next
    Cancel = False
End Sub

'''����O�ɏ����������N���A
Private Sub evtBk_BeforeClose(Cancel As Boolean)
    Dim sht As Worksheet
    For Each sht In evtBk.Sheets
        UnSetSelectionGriding sht
    Next
    Cancel = False
End Sub

