VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem CellHighlighter
Rem
Rem
Rem
Rem
Option Explicit

Private WithEvents App As Excel.Application
Attribute App.VB_VarHelpID = -1

Private excelWindow As Window

Private Sub Class_Initialize()
    Set App = Application
End Sub

'�C�x���g���t�b�N����I�u�W�F�N�g���w�肷��
'���b��łƂ��ăE�B���h�E�P�ʂŋL��
Public Function Init(obj As Object) As Object
    Const PROC_NAME = "CellViewer init"
    
    Select Case TypeName(obj)
        Case "Window": Set excelWindow = obj
        Case "Workbook": Set excelWindow = obj.Windows(1)
        Case "Worksheet": Set excelWindow = obj.Parent.Windows(1)
        '���̌^�ւ̑Ή��͕K�v�Ȃ���
        Case Else: Debug.Print PROC_NAME, "No Defined TypeName: " & TypeName(obj): Stop
    End Select

End Function

'�Z���̑I��ύX����A��ɋL�ڂ���Ă���Q�Ɛ�̃Z�����n�C���C�g����
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorBreak
    If Sh.Parent.Windows(1).Caption <> excelWindow.Caption Then Exit Sub
    If IsEmpty(Target.Value) Then Exit Sub
    On Error GoTo 0
    
    On Error Resume Next
        Dim refAdr As String: refAdr = Target.EntireRow.Cells(1, 1).Value
        Dim refRng As Range: Set refRng = GetRangeByFormula(refAdr)
        If refRng Is Nothing Then Debug.Print "address error : " & refAdr: Exit Sub
        Application.GoTo refRng
        If Err Then Exit Sub
        Call moveCenterRange(refRng)
    On Error GoTo 0
    
    Application.GoTo Target
    
    Call rangeHighlight(refRng, GLOBAL_HIGHLIGHT_NAME)
    
ErrorBreak:
End Sub
