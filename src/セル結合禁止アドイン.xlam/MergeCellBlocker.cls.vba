VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergeCellBlocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MergeBlcoker
'
'������u�b�N�̏㏑���ۑ������m���Z���̌������Ȃ����`�F�b�N����N���X
'
'�f�[�^�̓V�[�g����񎟌��z��œǂނ悤�ɏC���\��
'���b�Z�[�W�̏o�����͗v���P
'�������������B
'
Option Explicit

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1


Const ExMessage00 = "�Z���̌����� [num]�� �܂܂�Ă��܂��B"

Const ExMessage10 = "�y�x���z"
Const ExMessage11 = "�Z���̌����́A���Ȃ��̍�ƌ�����ቺ�����鋰�ꂪ����܂��B"
Const ExMessage12 = "�Z���̌������܂�Excel�t�@�C����z�z���邱�Ƃ́A�g�D�S�̂̋Ɩ�������ቺ�����鋰�ꂪ����܂��B"
Const ExMessage13 = "�Z���̌������܂�Excel�t�@�C����z�z���邱�ƂŁA����̐l��s�K�ɂ��鋰�ꂪ����܂��B"
Const ExMessage14 = "�Z���̌������܂�Excel�t�@�C����z�z���邱�ƂŁA���Ȃ�������̐l����ӂ߂��鋰�ꂪ����܂��B"
Const ExMessage19 = "����ł��ۑ����܂����H"

Const ExMessage20 = "�y��āz"
Const ExMessage21 = "�Z���̌������������邱�ƂŁA���ʂȍ�Ƃ��팸�ł��邩������܂���B"
Const ExMessage22 = "�Z���̌������������邱�Ƃ́A���Ȃ���Excel�X�L������Ɍq����܂��B"
Const ExMessage23 = "�Z���̌������������邱�ƂŁA�Г��ł̗F�D�֌W���ǂ��Ȃ邩������܂���B"
Const ExMessage24 = "�Z���̌������������邱�ƂŁA�C�ɂȂ邠�̎q���b�������Ă���邩���m��܂���B"
Const ExMessage29 = "�Z���̌��������ꂽ�ꏊ���m�F���܂����H"

Const ExMessage30 = "�y���߁z"
Const ExMessage31 = "�����������킸�ɂ������ƒ����񂩁`���I"
Const ExMessage39 = ""

Const OkMessage = "�����͊��S�ɋ쒀����܂���"

Property Get MessageTitle() As Collection
    Dim Col As Collection: Set Col = New Collection
    Col.Add ExMessage10
    Col.Add ExMessage20
    Col.Add ExMessage30
    Set MessageTitle = Col
End Property

Property Get MessageStyle() As Collection
    Dim Col As Collection: Set Col = New Collection
    Col.Add VbMsgBoxStyle.vbYesNo + VbMsgBoxStyle.vbExclamation
    Col.Add VbMsgBoxStyle.vbYesNo + VbMsgBoxStyle.vbInformation
    Col.Add VbMsgBoxStyle.vbOKOnly + VbMsgBoxStyle.vbCritical
    Set MessageStyle = Col
End Property

Property Get MessageData() As Collection
    Dim Col As Collection: Set Col = New Collection
    Col.Add Array(ExMessage11, ExMessage12, ExMessage13, ExMessage14)
    Col.Add Array(ExMessage21, ExMessage22, ExMessage23, ExMessage24)
    Col.Add Array(ExMessage31)
    Set MessageData = Col
End Property

Property Get MessageNextResult() As Collection
    Dim Col As Collection: Set Col = New Collection
    Col.Add VbMsgBoxResult.vbYes
    Col.Add VbMsgBoxResult.vbNo
    Col.Add VbMsgBoxResult.vbOK
    Set MessageNextResult = Col
End Property

Property Get MessageFooter() As Collection
    Dim Col As Collection: Set Col = New Collection
    Col.Add ExMessage19
    Col.Add ExMessage29
    Col.Add ExMessage39
    Set MessageFooter = Col
End Property

'�u�b�N�ۑ���
Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If CheckMergeCells(wb) Then
        Cancel = True
    End If
End Sub

'�����Z�������m���Ď���
'�ǂ��q�̂��߂Ɍ��������x���c�[�����N������
Private Function CheckMergeCells(wb As Workbook) As Boolean
'    If Not wb.Name Like "*.xls*" Then Exit Function
    If wb.IsAddin Then Exit Function
    
    Dim dic: Set dic = GetWorkbookMergeCellsDictionary(wb)
    If dic.Count = 0 Then
        Call removeHighlight(wb, GLOBAL_HIGHLIGHT_NAME)
'        wb.Windows(1).WindowState = xlMaximized
        MsgBox "�Z���̌����͂���܂���ł����B", vbOKOnly + vbInformation, APP_NAME
        Exit Function
    End If
    
    CheckMergeCells = True
    
    '���߂�܂Ń��b�Z�[�W��\��
    Dim i As Long
    For i = 1 To MessageTitle.Count
        Dim Item
        For Each Item In MessageData(i)
            If MsgBox(Item & vbLf & vbLf & MessageFooter(i), _
                MessageStyle(i), _
                MessageTitle(i) & " - " & Replace(ExMessage00, "[num]", dic.Count) _
                ) <> MessageNextResult(i) Then
                GoTo BreakForFor
            End If
        Next
    Next
BreakForFor:

    '�Z���̌����̉������J�n
    Call ViewMergeCells(wb)
    MsgBox "�Z���̌����̉����c�[�����N�����܂����B", vbOKOnly + vbInformation, APP_NAME
    
End Function

Rem �I�u�W�F�N�g�̍쐬
Public Function Init(pApp As Excel.Application) As MergeCellBlocker
    If Me Is MergeCellBlocker Then
        With New MergeCellBlocker
            Set Init = .Init(pApp)
        End With
        Exit Function
    End If
    Set Init = Me
    Set app = pApp
End Function

'���������_���ɂ��āA�͂��@�Ɓ@�������@�����ւ���
