Attribute VB_Name = "FuncExcel_Partial"
'ちゅんちゅんライブラリのFuncExcelから切り出した機能
Option Explicit

'指定した名前のShapeが一つ以上定義されているか確認する。
'但し同一名のシェイプは作成できるので基本的には下記を使用する。
Public Function ShapeExists(ShapeName As String, _
                            Optional ws As Worksheet, _
                            Optional wb As Workbook) As Boolean
    Dim sps As Collection
    Set sps = ShapesFill(ShapeName, ws, wb)
    ShapeExists = (sps.Count > 0)
End Function

'指定した名前に一致するシェイプを抽出する。
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
