VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PassForm 
   Caption         =   "Login"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   OleObjectBlob   =   "PassForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PassInputted As String
Public ButtonInputted As Integer

Private Sub LoginPassFormQuit(ret As Integer)
    Me.PassInputted = Me.TextBox1.Value
    Me.ButtonInputted = ret
    
    Me.Hide
End Sub

Private Sub CommandButton1_Click()
    Call LoginPassFormQuit(1)
End Sub

Private Sub CommandButton2_Click()
    Call LoginPassFormQuit(2)
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
    If closemode = vbFormControlMenu Then
        Call LoginPassFormQuit(0)
    End If
End Sub
