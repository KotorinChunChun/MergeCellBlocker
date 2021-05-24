Attribute VB_Name = "ProcMergeCells"
Option Explicit

Public Const GLOBAL_HIGHLIGHT_NAME = "MergeCellHighlight"

'���b�Z�[�W�ҏW��ʂ�\��


'���b�Z�[�W�ҏW��ʂ��B��

'�A�N�e�B�u�u�b�N�Ńe�X�g���J�n
Sub Test_ViewMergeCells(): Call ViewMergeCells(ActiveWorkbook): End Sub


'�����Z����V�K�u�b�N�Ɉꗗ�ŕ\��
Sub ViewMergeCells(targetWb As Workbook)

    Dim dic As Dictionary: Set dic = GetWorkbookMergeCellsDictionary(targetWb)
    If dic.Count = 0 Then Exit Sub
    
    '�������ꂽ�ꏊ�̈ꗗ��\��
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
    
    '�E�B���h�E���̔������@�{�c��
'    Dim mainWindowWidth As Long
'    mainWindowWidth = targetWb.Windows(1).Width - viewWs.Columns(1).Width - viewWs.Columns(2).Width
'    Debug.Print mainWindowWidth
    
    '�E�B���h�E�����E�ɕ��ׂĕ\��
'    ExcelWindowArrange Array(targetWb.Windows(1), viewWb.Windows(1)), xlVertical
    targetWb.Windows(1).WindowState = xlMaximized
    ExcelWindowArrange Array(targetWb, viewWb), xlVertical
    
    '�r���[���[�̊Ď����J�n
    Call CellHighlightStartWs(viewWs)
    
End Sub

'�u�b�N�̌������ꂽ�Z��(�擪�Z��)�����X�g�A�b�v����֐�
' Dictionary(�Q�Ɛ���, �l)
' �Q�Ɛ�����Excel�Ɠ����d�l�F'[BookName]SheetName'!Address
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

'�n�C���C�g�̃e�X�g
Sub Test_rangeHighlight()
    Call rangeHighlight(Range("D13"), "test")
    Stop
    Call rangeHighlight(Range("G23"), "test")
End Sub

'�w�肵���Z�����n�C���C�g����
Public Sub rangeHighlight(rng As Range, obj_name As String)
    Const pointMargin = 5
    
    '�͂��}�`������ or �ė��p
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
    
    '�������l�������Z���̗̈��}�`�ň͂�
    shp.Left = rng.Left - pointMargin
    shp.Top = rng.Top - pointMargin
    shp.Width = rng.Offset(1, 1).Left - rng.Left + pointMargin * 2
    shp.Height = rng.Offset(1, 1).Top - rng.Top + pointMargin * 2
    
End Sub

'�w�薼�̂̃n�C���C�g�p�̐}�`���폜����
Sub removeHighlight(wb As Workbook, obj_name As String)
    On Error Resume Next
    Dim shps: Set shps = kccFuncExcel.ShapesFill(obj_name, wb:=wb)
    Dim shp
    For Each shp In shps
        shp.Delete
    Next
End Sub

'�w�肵���Z������ʂ̒����ɗ���悤�ɃX�N���[������
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

'����̃G�N�Z���E�B���h�E��������ׂĕ\������
Sub ExcelWindowArrange(targets, arrange_style As XlOrientation)
    Const PROC_NAME = "ExcelWindowArrange"

    Dim win As Window
    Dim obj As Variant
    
    'targets����́F����Ώۂ�Window�R���N�V����������
    Dim arrangeWindows As Collection: Set arrangeWindows = New Collection
    For Each obj In targets
        Select Case TypeName(obj)
            Case "Window": arrangeWindows.Add obj
            Case "Workbook": arrangeWindows.Add obj.Windows(1)
            Case "Worksheet": arrangeWindows.Add obj.Parent.Windows(1)
            '���̌^�ւ̑Ή��͕K�v�Ȃ���
            Case Else: Debug.Print PROC_NAME, "No Defined TypeName: " & TypeName(obj): Stop
        End Select
    Next
    
    If arrangeWindows.Count < 2 Then
        Debug.Print PROC_NAME, "targets������Ȃ�"
'        Err.Raise 9999, PROC_NAME, "targets������Ȃ�"
        Exit Sub
    End If
    
    '����ΏۊO�Ƃ��鑋��Window�R���N�V����������
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
    
    '��\���ɂ��Đ��񂳂�Ȃ��悤�ɂ���@�i�ʈā@win.WindowState = xlMinimized�j
    For Each win In withoutWindows: win.Visible = False: Next
    
    '2�Ԗڈȍ~�̃E�B���h�E��擪�Ɠ����f�X�N�g�b�v�ցB
    arrangeWindows(1).WindowState = xlNormal    '�ő厞�̍��W�͐ݒ�o���Ȃ�
'    Debug.Print arrangeWindows(1).Caption, arrangeWindows(1).Left, arrangeWindows(1).Top
    For Each win In arrangeWindows
        win.WindowState = xlNormal
        win.Left = arrangeWindows(1).Left
        win.Top = arrangeWindows(1).Top
    Next
    
    '���ׂĕ\��
    Windows.Arrange ArrangeStyle:=arrange_style
    
    '��\���ɂ��Ă����������ĕ\���@�i�ʈā@win.WindowState = xlNormal�j
    For Each win In withoutWindows: win.Visible = True: Next
    
    '����Ώۂ̃E�B���h�E��O�ʂցB�擪�����A�N�e�B�u��
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
