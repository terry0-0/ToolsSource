VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EvtSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Public WithEvents evtSht As Worksheet
Attribute evtSht.VB_VarHelpID = -1

 
Private Sub evtSht_SelectionChange(ByVal Target As Range)
    evtSht.Application.ScreenUpdating = True
    evtSht.Calculate
End Sub

