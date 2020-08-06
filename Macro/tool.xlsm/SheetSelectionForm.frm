VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetSelectionForm 
   Caption         =   "Select Sheet"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11220
   OleObjectBlob   =   "SheetSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SheetSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim resultArr As Collection

Public Sub setResultArr(arr)
    Set resultArr = arr
    Me.ListBox1.Clear
    Dim sht
    For Each sht In resultArr
        Me.ListBox1.AddItem sht.Name
    Next
    
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ShowSheet ListBox1.Text
    Me.Hide
End Sub

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 13 'Enter
            ShowSheet ListBox1.Text
            Me.Hide
        Case 27 'Esc
            Me.Hide
    End Select
    
End Sub

Private Sub ShowSheet(str)
    Dim sht
    For Each sht In resultArr
        If sht.Name = str Then
            sht.Parent.Activate
            sht.Activate
        End If
    Next
    Me.Hide
End Sub
