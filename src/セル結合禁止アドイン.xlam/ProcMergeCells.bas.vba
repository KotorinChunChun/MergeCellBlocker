Attribute VB_Name = "ProcMergeCells"
Option Explicit

Public Const GLOBAL_HIGHLIGHT_NAME = "MergeCellHighlight"

'メッセージ編集画面を表示


'メッセージ編集画面を隠す

'アクティブブックでテストを開始
Sub Test_ViewMergeCells(): Call ViewMergeCells(ActiveWorkbook): End Sub


'結合セルを新規ブックに一覧で表示
Sub ViewMergeCells(targetWb As Workbook)

    Dim dic As Dictionary: Set dic = GetWorkbookMergeCellsDictionary(targetWb)
    If dic.Count = 0 Then Exit Sub
    
    '結合された場所の一覧を表示
    Dim viewWb As Workbook: Set viewWb = Workbooks.Add
    Dim viewWs As Worksheet: Set viewWs = viewWb.Worksheets(1)
    
    Dim data As Variant
    ReDim data(1 To 2)
    data(1) = dic.Keys
    data(2) = dic.Items
    data = WorksheetFunction.Transpose(data)
    
    viewWs.Columns("A:B").NumberFormatLocal = "@"
    viewWs.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value = data
    viewWs.Columns.AutoFit
    
    'ウィンドウ幅の微調整　ボツ案
'    Dim mainWindowWidth As Long
'    mainWindowWidth = targetWb.Windows(1).Width - viewWs.Columns(1).Width - viewWs.Columns(2).Width
'    Debug.Print mainWindowWidth
    
    'ウィンドウを左右に並べて表示
'    ExcelWindowArrange Array(targetWb.Windows(1), viewWb.Windows(1)), xlVertical
    targetWb.Windows(1).WindowState = xlMaximized
    ExcelWindowArrange Array(targetWb, viewWb), xlVertical
    
    'ビューワーの監視を開始
    Call CellHighlightStartWs(viewWs)
    
End Sub

'ブックの結合されたセル(先頭セル)をリストアップする関数
' Dictionary(参照数式, 値)
' 参照数式はExcelと同じ仕様：'[BookName]SheetName'!Address
Public Function GetWorkbookMergeCellsDictionary(wb As Workbook) As Object
        
    Dim ws As Worksheet
    Dim rng As Range
    Dim ret As Object: Set ret = CreateObject("Scripting.Dictionary")
    
    For Each ws In wb.Worksheets
        For Each rng In ws.UsedRange
            If rng.MergeCells Then
                If rng.MergeArea.Item(1, 1).Address = rng.Item(1, 1).Address Then
                    ret.Add "='[" & wb.Name & "]" & ws.Name & "'!" & rng.Address(False, False), rng.Item(1, 1).Value
                End If
            End If
        Next
    Next
    
    Set GetWorkbookMergeCellsDictionary = ret
    
End Function

'ハイライトのテスト
Sub Test_rangeHighlight()
    Call rangeHighlight(Range("D13"), "test")
    Stop
    Call rangeHighlight(Range("G23"), "test")
End Sub

'指定したセルをハイライトする
Public Sub rangeHighlight(rng As Range, obj_name As String)
    Const pointMargin = 5
    
    '囲い図形を準備 or 再利用
    Dim shp As Shape
    Dim shps: Set shps = kccFuncExcel.ShapesFill(obj_name, rng.Worksheet)
    
    If shps.Count = 0 Then
        Set shp = rng.Worksheet.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        shp.Name = obj_name
        With shp.Fill
            .Visible = msoFalse
            .Transparency = 0
            .Solid
        End With
        
        With shp.Line
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(255, 0, 0)
            .Weight = 6
        End With
    Else
        Set shp = shps(1)
    End If
    
    '結合を考慮したセルの領域を図形で囲う
    shp.Left = rng.Left - pointMargin
    shp.Top = rng.Top - pointMargin
    shp.Width = rng.Offset(1, 1).Left - rng.Left + pointMargin * 2
    shp.Height = rng.Offset(1, 1).Top - rng.Top + pointMargin * 2
    
End Sub

'指定名称のハイライト用の図形を削除する
Sub removeHighlight(wb As Workbook, obj_name As String)
    On Error Resume Next
    Dim shps: Set shps = kccFuncExcel.ShapesFill(obj_name, wb:=wb)
    Dim shp
    For Each shp In shps
        shp.Delete
    Next
End Sub

'指定したセルが画面の中央に来るようにスクロールする
Sub moveCenterRange(rng As Range)
    With rng.Worksheet.Parent.Windows(1)
        Dim c As Long, r As Long
        c = rng.Column - .VisibleRange.Cells.Columns.Count / 2
        r = rng.Row - .VisibleRange.Cells.Rows.Count / 2
        If c < 1 Then c = 1
        If r < 1 Then r = 1
        .ScrollColumn = c
        .ScrollRow = r
    End With
End Sub

'特定のエクセルウィンドウだけを並べて表示する
Sub ExcelWindowArrange(targets, arrange_style As XlOrientation)
    Const PROC_NAME = "ExcelWindowArrange"

    Dim win As Window
    Dim obj As Variant
    
    'targetsを解析：整列対象のWindowコレクションを準備
    Dim arrangeWindows As Collection: Set arrangeWindows = New Collection
    For Each obj In targets
        Select Case TypeName(obj)
            Case "Window": arrangeWindows.Add obj
            Case "Workbook": arrangeWindows.Add obj.Windows(1)
            Case "Worksheet": arrangeWindows.Add obj.Parent.Windows(1)
            '他の型への対応は必要なら作る
            Case Else: Debug.Print PROC_NAME, "No Defined TypeName: " & TypeName(obj): Stop
        End Select
    Next
    
    If arrangeWindows.Count < 2 Then
        Debug.Print PROC_NAME, "targetsが足りない"
'        Err.Raise 9999, PROC_NAME, "targetsが足りない"
        Exit Sub
    End If
    
    '整列対象外とする窓のWindowコレクションを準備
    Dim withoutWindows As Collection: Set withoutWindows = New Collection
    For Each win In Application.Windows
        If win.Visible Then
            Dim arrWin As Window
            For Each arrWin In arrangeWindows
                If win.Caption = arrWin.Caption Then GoTo ContinueFor
            Next
            withoutWindows.Add win
        End If
ContinueFor:
    Next
    
    '--------------------------------------------------
    
    '非表示にして整列されないようにする　（別案　win.WindowState = xlMinimized）
    For Each win In withoutWindows: win.Visible = False: Next
    
    '2番目以降のウィンドウを先頭と同じデスクトップへ。
    arrangeWindows(1).WindowState = xlNormal    '最大時の座標は設定出来ない
'    Debug.Print arrangeWindows(1).Caption, arrangeWindows(1).Left, arrangeWindows(1).Top
    For Each win In arrangeWindows
        win.WindowState = xlNormal
        win.Left = arrangeWindows(1).Left
        win.Top = arrangeWindows(1).Top
    Next
    
    '並べて表示
    Windows.Arrange ArrangeStyle:=arrange_style
    
    '非表示にしておいた窓を再表示　（別案　win.WindowState = xlNormal）
    For Each win In withoutWindows: win.Visible = True: Next
    
    '整列対象のウィンドウを前面へ。先頭窓をアクティブへ
    For Each win In arrangeWindows: win.Activate: Next
    arrangeWindows(1).Activate

End Sub

Sub SubMergeCellCreater(IsStop As Boolean)
    Static mcc As MergeCellCreater
    
    If IsStop Then
        Set mcc = Nothing
        Exit Sub
    End If
    
    Set mcc = MergeCellCreater.Init(Application)
End Sub
