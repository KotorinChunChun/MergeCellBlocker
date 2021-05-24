VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not Me.Saved Then
        If MsgBox(APP_NAME & "は変更されています。保存してから終了しますか？", vbYesNo) = vbYes Then
            Me.Save
        End If
    End If
End Sub

'アドインが起動されたらクラスを生成
Private Sub Workbook_Open()
    Call AddinStart
End Sub
