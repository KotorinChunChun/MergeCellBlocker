VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncExcel_Partial
Rem
Rem  @description   Excelを拡張する汎用的な関数
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
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

Rem 指定Rangeを同一値でセルを結合
Rem targetRange      : 結合したい範囲(SelectionRange等）
Rem CanRowMerge   : 行結合を認める
Rem CanColMerge   : 列結合を認める
Rem CanEmptyMerge : 空白行の結合を認める
Rem MergePriority_Sum0Down1Right2   : 結合優先度（合計数→行優先→列優先の3択）
Function RangeMergeByValue( _
                    ByVal targetRange As Range, _
                    Optional CanRowMerge As Boolean = True, _
                    Optional CanColMerge As Boolean = True, _
                    Optional CanEmptyMerge As Boolean = True, _
                    Optional MergePriority_Sum0Down1Right2 As Long) As Range
    
    Dim ur As Excel.Range: Set ur = targetRange.Worksheet.UsedRange
    Set ur = ur.Resize(ur.Rows.Count + IIf(ur.Rows.CountLarge < ur.Worksheet.Rows.CountLarge, 1, 0), _
                        ur.Columns.Count + IIf(ur.Columns.CountLarge < ur.Worksheet.Columns.CountLarge, 1, 0))
    Set targetRange = Intersect(targetRange, ur)
    If targetRange Is Nothing Then Exit Function
    
    If targetRange.Areas.Count > 1 Then
        Dim rngArea As Range
        For Each rngArea In targetRange.Areas
            Call RangeMergeByValue(rngArea, CanRowMerge:=CanRowMerge, CanColMerge:=CanColMerge, CanEmptyMerge:=CanEmptyMerge, MergePriority_Sum0Down1Right2:=MergePriority_Sum0Down1Right2)
        Next
        Exit Function
    End If
    
    Dim r As Range
    Dim i As Long, j As Long, k As Long
    Dim v As Variant
    Dim StackR As Range
    Dim MaxRow As Long
    Dim MaxCol As Long
    MaxRow = targetRange.Rows.Count
    MaxCol = targetRange.Columns.Count
    
    Application.DisplayAlerts = False
    Set StackR = targetRange.Cells(1, 1).MergeArea
    v = StackR.Cells(1, 1).Value
    For j = 1 To MaxCol
        For i = 1 To MaxRow
            Set r = targetRange.Cells(i, j).MergeArea.Cells(1, 1)
            
            '直前のセル有、値一致、セル結合準備
            If Not StackR Is Nothing Then
                If v = r.Value Then
                    Set StackR = Union(StackR, r)
                End If
            End If
            
            '値不一致 or 最終行
            If v <> r.Value Or i = MaxRow Then
                '結合候補有：結合実施
                If Not StackR Is Nothing Then
                    If StackR.Count > 1 Then
                        StackR.Merge
                    End If
                End If
                '結合先頭セルをセット
                Set StackR = r
                v = r.Value
            End If
        Next
    Next
    
    Set StackR = targetRange.Cells(1, 1).MergeArea
    v = StackR.Cells(1, 1).Value
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            Set r = targetRange.Cells(i, j).MergeArea.Cells(1, 1)
            
            '直前のセル有、値一致、セル結合準備
            If Not StackR Is Nothing Then
                If v = r.Value And StackR.Cells(1, 1).MergeArea.Rows.Count = r.MergeArea.Rows.Count Then
                    Set StackR = Union(StackR, r.MergeArea)
                End If
            End If
            
            '値不一致 or 最終列
            If v <> r.Value Or j = MaxCol Then
                '結合候補有：結合実施
                If Not StackR Is Nothing Then
                    If StackR.Count > 1 Then
                        StackR.Merge
                    End If
                End If
                '結合先頭セルをセット
                Set StackR = r.MergeArea
                v = r.Value
            End If
            
        Next
        Set StackR = Nothing
        v = ""
    Next
    Application.DisplayAlerts = True
    
End Function

Rem セル結合を解除して同一値で埋める
Rem  @param targetRange    : 結合範囲(SelectionRange等）
Rem  @param CanRowMerge : 行結合を認める
Rem  @param CanColMerge : 列結合を認める
Function RangeUnMerge(targetRange As Range, _
                        Optional CanRowMerge As Boolean = True, _
                        Optional CanColMerge As Boolean = True) As Range
                    
    Set targetRange = Intersect(targetRange, targetRange.Worksheet.UsedRange)
    If targetRange Is Nothing Then Exit Function

    Dim area As Range
    Dim rng As Range
    Dim adr As String
    Dim rgs As Collection: Set rgs = New Collection
    
    '結合されたセルで、左上のRangeみをリストアップ
    For Each area In targetRange.Areas
        'Debug.Print "Area : " & Area.Address
        For Each rng In area.Cells
            If rng.MergeCells And rng.Address = rng.MergeArea(1).Address Then
                rgs.Add rng
                'Debug.Print rng.Address & " <<< " & rng.MergeArea.Address
            End If
        Next
    Next
    
    '結合セルを分解
    Dim v As Variant
    For Each rng In rgs
        Dim bdhWeight, bdwWeight
        bdhWeight = rng.Borders(xlEdgeTop).Weight
        bdwWeight = rng.Borders(xlEdgeLeft).Weight
        v = rng(1, 1).Value
        adr = rng.MergeArea.Address
        rng.MergeArea.MergeCells = False
        With rng.Parent.Range(adr)
            .Value = v
            '罫線を復元
            .Borders(xlEdgeTop).Weight = bdhWeight
            .Borders(xlInsideHorizontal).Weight = bdhWeight
            .Borders(xlEdgeBottom).Weight = bdhWeight
            
            .Borders(xlEdgeLeft).Weight = bdwWeight
            .Borders(xlInsideVertical).Weight = bdwWeight
            .Borders(xlEdgeRight).Weight = bdwWeight
        End With
    Next
    
End Function

Rem 同一値のセルを横に結合または縦に結合
Function RangeMergeRightDownByValue(targetRange As Range, Down1Right2)
    Set targetRange = Intersect(targetRange, targetRange.Worksheet.UsedRange)
    If targetRange Is Nothing Then Exit Function
    
    If targetRange.Areas.Count > 1 Then
        Dim rngArea As Range
        For Each rngArea In targetRange.Areas
            Call RangeMergeRightDownByValue(rngArea, Down1Right2)
        Next
        Exit Function
    End If
    
    Application.DisplayAlerts = False
    Dim rngLine As Range
    Dim rngCell As Range
    Dim mergeRange As Range
    For Each rngLine In IIf(Down1Right2 = 1, targetRange.Columns, targetRange.Rows)
        Set mergeRange = Nothing
        For Each rngCell In rngLine.Cells
            
'            If Not mergeRange Is Nothing Then
'                Debug.Print mergeRange.Address(False, False), mergeRange.Columns.CountLarge, mergeRange.Rows.CountLarge
'                Debug.Print rngCell.Address(False, False), rngCell.MergeArea.Columns.CountLarge, rngCell.MergeArea.Rows.CountLarge
'                Excel.Range(mergeRange.Address & "," & rngCell.Address).Select
'                Stop
'            End If
            
            If mergeRange Is Nothing Then
                Set mergeRange = rngCell.MergeArea
            ElseIf mergeRange(1).MergeCells Then
                If mergeRange(1).Value = rngCell.Value And _
                        ((Down1Right2 = 1 And mergeRange.Columns.CountLarge = rngCell.MergeArea.Columns.CountLarge) Or _
                        (Down1Right2 = 2 And mergeRange.Rows.CountLarge = rngCell.MergeArea.Rows.CountLarge)) Then
                    
                    Set mergeRange = Excel.Union(mergeRange, rngCell.MergeArea)
'                    mergeRange.Merge
'                    mergeRange.Select
                Else
                    mergeRange.Merge
                    mergeRange.Select
                    Set mergeRange = rngCell.MergeArea
                End If
            ElseIf mergeRange(1).Value = rngCell.Value And _
                (mergeRange(1).MergeArea.Columns.CountLarge = rngCell.MergeArea.Columns.CountLarge Or _
                mergeRange(1).MergeArea.Rows.CountLarge = rngCell.MergeArea.Rows.CountLarge) Then
                
                Set mergeRange = Excel.Union(mergeRange, rngCell.MergeArea)
'                mergeRange.Merge
'                mergeRange.Select
            Else
                mergeRange.Merge
                Set mergeRange = rngCell.MergeArea
            End If
        Next
        If Not mergeRange Is Nothing Then mergeRange.Merge
    Next
    Application.DisplayAlerts = True
End Function
