Attribute VB_Name = "FuncExcel_Partial"
'����񂿂�񃉃C�u������FuncExcel����؂�o�����@�\
Option Explicit

'�w�肵�����O��Shape����ȏ��`����Ă��邩�m�F����B
'�A�����ꖼ�̃V�F�C�v�͍쐬�ł���̂Ŋ�{�I�ɂ͉��L���g�p����B
Public Function ShapeExists(ShapeName As String, _
                            Optional ws As Worksheet, _
                            Optional wb As Workbook) As Boolean
    Dim sps As Collection
    Set sps = ShapesFill(ShapeName, ws, wb)
    ShapeExists = (sps.Count > 0)
End Function

'�w�肵�����O�Ɉ�v����V�F�C�v�𒊏o����B
Public Function ShapesFill(ShapeName As String, _
                            Optional ws As Worksheet, _
                            Optional wb As Workbook) As Collection
    Dim sp As Shape
    Dim CL As New Collection
    
    If ws Is Nothing Then
        If wb Is Nothing Then
            Set ws = ActiveWorkbook.ActiveSheet
        Else
            Set ws = wb.ActiveSheet
        End If
    End If
    
    For Each sp In ws.Shapes
        If sp.Name = ShapeName Then
            CL.Add sp
        End If
    Next
    Set ShapesFill = CL
End Function
