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

'����ł��h��풓�}�N��
'��{�I�Ƀv���W�F�N�g���Ⴆ�΃��Z�b�g���Ă����ȂȂ��̂ŕK�v�͂Ȃ��B
'�u���[�N�|�C���g�Œ�~���������o���ăG���[��f�����߁AVBAer�ɂƂ��Ă͎ז��ł��������B

'Private Property Get rngSW() As Range: Set rngSW = Sheet1.Cells(1, 1): End Property
'
'Private Property Get EnableMergeBlocker(): EnableMergeBlocker = rngSW.Value: End Property
'Private Property Let EnableMergeBlocker(tf): rngSW.Value = tf: End Property
'
'Private Sub NonDeadMergeBlocker()
'    Const PROC_NAME = "ThisWorkbook.NonDeadMergeBlocker"
'    Const WAIT_SECOND = 5
'    Static nextTime As Date
'    Debug.Print Format(Now, "hh:mm:ss"), PROC_NAME, ;
'
'    If EnableMergeBlocker Then
'        Debug.Print "enable ", ;
'
'        nextTime = Now() + TimeSerial(0, 0, WAIT_SECOND)
'        Debug.Print "next:" & Format(nextTime, "hh:mm:ss"), ;
'
'        If instMergeBlocker Is Nothing Then
'            Set instMergeBlocker = New MergeBlocker
'            Debug.Print "new MergeBlocker", ;
'        End If
'
'        Application.OnTime nextTime, PROC_NAME
'        Debug.Print
'    Else
'        Debug.Print "disable"
'    End If
'End Sub
'
'Public Sub MergeBlockerStart()
'    If Not EnableMergeBlocker Then
'        EnableMergeBlocker = True
'        Call NonDeadMergeBlocker
'    End If
'End Sub
'Public Sub MergeBlockerStop()
'    EnableMergeBlocker = False
'End Sub

'�A�h�C�����N�����ꂽ��N���X�𐶐�
Private Sub Workbook_Open()
    Call MonitorStart
End Sub
