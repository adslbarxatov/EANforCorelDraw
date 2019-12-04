VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputValueForm 
   Caption         =   "Укажите значение штрихкода"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   OleObjectBlob   =   "InputValueForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InputValueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ValueSelected As Boolean

Private Sub BAccept_Click()
    If IsNumeric(BarcodeValue.Text) = False Then
       Call MsgBox("Введённое значение не является числом", vbOKOnly + vbExclamation)
       Exit Sub
    End If

    ValueSelected = True
    Call Me.Hide
End Sub

Private Sub UserForm_Initialize()
    ValueSelected = False
End Sub

