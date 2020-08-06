VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "InputForm"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private targetRng As Range

Public Sub ShowInputFormArr(rng As Range, arr)
    Set targetRng = rng
    Me.ListBox1.Clear
    Dim i
    For i = 0 To UBound(arr)
        If Trim(arr(i)) <> Empty Then
        Me.ListBox1.AddItem arr(i)
            If rng.Cells(1, 1).Value = arr(i) Then
                Me.ListBox1.ListIndex = i
            End If
        End If
    Next
    Me.Show
End Sub

Public Sub ShowInputFormList(rng As Range, list)
    Set targetRng = rng
    Me.ListBox1.Clear
    Dim i
    Dim obj
    For Each obj In list
        If Trim(CStr(obj)) <> Empty Then
            Me.ListBox1.AddItem obj
            If rng.Cells(1, 1).Value = obj Then
                Me.ListBox1.ListIndex = i
            End If
            i = i + 1
        End If
    Next
    Me.Show
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ChangeValue ListBox1.Text
    Me.Hide
End Sub

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 13 'Enter
            ChangeValue ListBox1.Text
            Me.Hide
        Case 27 'Esc
            Me.Hide
    End Select
    
End Sub


Private Sub ChangeValue(str)
    Dim rng As Range
    For Each rng In targetRng
        If rng.Rows.Hidden = False And rng.Columns.Hidden = False Then
            rng.Value = str
        End If
    Next
End Sub
