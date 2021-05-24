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
'あらゆるブックの上書き保存を検知しセルの結合がないかチェックするクラス
'
'データはシートから二次元配列で読むように修正予定
'メッセージの出し方は要改善
'書き直したい。
'
Option Explicit

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1


Const ExMessage00 = "セルの結合が [num]件 含まれています。"

Const ExMessage10 = "【警告】"
Const ExMessage11 = "セルの結合は、あなたの作業効率を低下させる恐れがあります。"
Const ExMessage12 = "セルの結合を含むExcelファイルを配布することは、組織全体の業務効率を低下させる恐れがあります。"
Const ExMessage13 = "セルの結合を含むExcelファイルを配布することで、周りの人を不幸にする恐れがあります。"
Const ExMessage14 = "セルの結合を含むExcelファイルを配布することで、あなたが周りの人から責められる恐れがあります。"
Const ExMessage19 = "それでも保存しますか？"

Const ExMessage20 = "【提案】"
Const ExMessage21 = "セルの結合を解除することで、無駄な作業が削減できるかもしれません。"
Const ExMessage22 = "セルの結合を解除することは、あなたのExcelスキル向上に繋がります。"
Const ExMessage23 = "セルの結合を解除することで、社内での友好関係が良くなるかもしれません。"
Const ExMessage24 = "セルの結合を解除することで、気になるあの子が話しかけてくれるかも知れません。"
Const ExMessage29 = "セルの結合がされた場所を確認しますか？"

Const ExMessage30 = "【命令】"
Const ExMessage31 = "ぐだぐだ言わずにさっさと直さんか～い！"
Const ExMessage39 = ""

Const OkMessage = "結合は完全に駆逐されました"

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

'ブック保存時
Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If CheckMergeCells(wb) Then
        Cancel = True
    End If
End Sub

'結合セルを検知して叱る
'良い子のために結合解除支援ツールを起動する
Private Function CheckMergeCells(wb As Workbook) As Boolean
'    If Not wb.Name Like "*.xls*" Then Exit Function
    If wb.IsAddin Then Exit Function
    
    Dim dic: Set dic = GetWorkbookMergeCellsDictionary(wb)
    If dic.Count = 0 Then
        Call removeHighlight(wb, GLOBAL_HIGHLIGHT_NAME)
'        wb.Windows(1).WindowState = xlMaximized
        MsgBox "セルの結合はありませんでした。", vbOKOnly + vbInformation, APP_NAME
        Exit Function
    End If
    
    CheckMergeCells = True
    
    '諦めるまでメッセージを表示
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

    'セルの結合の解消を開始
    Call ViewMergeCells(wb)
    MsgBox "セルの結合の解消ツールを起動しました。", vbOKOnly + vbInformation, APP_NAME
    
End Function

Rem オブジェクトの作成
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

'問題をランダムにして、はい　と　いいえ　を入れ替える
