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
        If MsgBox(APP_NAME & "�͕ύX����Ă��܂��B�ۑ����Ă���I�����܂����H", vbYesNo) = vbYes Then
            Me.Save
        End If
    End If
End Sub

'�A�h�C�����N�����ꂽ��N���X�𐶐�
Private Sub Workbook_Open()
    Call AddinStart
End Sub
