Attribute VB_Name = "AppMain"
Rem
Rem @appname MergeCellBlocker - �Z�������֎~�A�h�C��
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/15 ����Łi���p���F���j
Rem    2020/02/19 �C���Łi�Ƃ肠�������p�j
Rem    2020/02/28 Git���J
Rem    2021/05/24 �Z�������x���ǉ�
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "�Z�������֎~�A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.13"
Public Const APP_UPDATE = "2021/05/24"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/merge_cell_blocker"

Public instMergeBlocker As MergeBlocker
Public instCellHighlighter As CellHighlighter

'--------------------------------------------------
'�A�h�C�����s��
Sub AddinStart()
    Call MonitorStart
    MsgBox "�Z���̌����͐�΂ɂ�邵�܂ւ�Ł`�`�`�I", _
                vbExclamation + vbOKOnly, ThisWorkbook.Name
End Sub

'�A�h�C���ꎞ��~��
Sub AddinStop(): Call MonitorStop: End Sub

'�A�h�C���ݒ�\��
Sub AddinConfig(): Call SettingForm.Show: End Sub

'�A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'�A�h�C�����~�߂������Ɏg���v���V�[�W��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

'������
Sub MergeSearch(): MsgBox "Search": End Sub

Rem �Z���������������Ēl�Ŗ��߂�
Sub MergeBreak():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeUnMerge(Selection)
End Sub

Rem �c�����ɓ����l���Z������
Sub MergeDown():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeRightDownByValue(Selection, 1)
'    Call kccFuncExcel.RangeMergeByValue(Selection, CanRowMerge:=False, CanColMerge:=True)
End Sub

Rem �������ɓ����l���Z������
Sub MergeRight()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeRightDownByValue(Selection, 2)
'    Call kccFuncExcel.RangeMergeByValue(Selection, CanRowMerge:=True, CanColMerge:=False)
End Sub

Rem �K���ȕ����փZ������
Sub MergeAuto():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeByValue(Selection, True, True, False, 0)
End Sub

Sub MergePrint(): MsgBox "Print": End Sub
'--------------------------------------------------

'�Ď��J�n
'Workbook_Open����Ă΂��
'���u�b�N�̏㏑���ۑ������m���邽�߂Ɏg�p�����
Sub MonitorStart(): Set instMergeBlocker = New MergeBlocker: End Sub

'�Ď���~
Sub MonitorStop(): Set instMergeBlocker = Nothing: End Sub

'�Z�������x��������ON
Sub Start_MergeCellCreater()
    Call SubMergeCellCreater(False)
End Sub

'�Z�������x��������OFF
Sub Stop_MergeCellCreater()
    Call SubMergeCellCreater(True)
End Sub

'�����Z���ꗗ�\����ɌĂяo���v���V�[�W��
'���u�b�N�̃Z���I�������m�����邽�߂Ɏg�p�����
Sub CellHighlightStart(): Call CellHighlightStartWs(ActiveSheet): End Sub
Sub CellHighlightStartWs(Optional ws As Worksheet)
    Set instCellHighlighter = New CellHighlighter
    instCellHighlighter.Init ws
End Sub
