VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SettingForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " - 設定"
End Sub

