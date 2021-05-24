Attribute VB_Name = "AppMain"
Rem
Rem @appname MergeCellBlocker - セル結合禁止アドイン
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/15 初回版（実用性皆無）
Rem    2020/02/19 修正版（とりあえず実用可）
Rem    2020/02/28 Git公開
Rem    2021/05/24 セル結合支援追加
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "セル結合禁止アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.13"
Public Const APP_UPDATE = "2021/05/24"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/merge_cell_blocker"

Public instMergeBlocker As MergeBlocker
Public instCellHighlighter As CellHighlighter

'--------------------------------------------------
'アドイン実行時
Sub AddinStart()
    Call MonitorStart
    MsgBox "セルの結合は絶対にゆるしまへんで～～～！", _
                vbExclamation + vbOKOnly, ThisWorkbook.Name
End Sub

'アドイン一時停止時
Sub AddinStop(): Call MonitorStop: End Sub

'アドイン設定表示
Sub AddinConfig(): Call SettingForm.Show: End Sub

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'アドインを止めたい時に使うプロシージャ
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

'未実装
Sub MergeSearch(): MsgBox "Search": End Sub

Rem セル結合を解除して値で埋める
Sub MergeBreak():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeUnMerge(Selection)
End Sub

Rem 縦方向に同じ値をセル結合
Sub MergeDown():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeRightDownByValue(Selection, 1)
'    Call kccFuncExcel.RangeMergeByValue(Selection, CanRowMerge:=False, CanColMerge:=True)
End Sub

Rem 横方向に同じ値をセル結合
Sub MergeRight()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeRightDownByValue(Selection, 2)
'    Call kccFuncExcel.RangeMergeByValue(Selection, CanRowMerge:=True, CanColMerge:=False)
End Sub

Rem 適当な方向へセル結合
Sub MergeAuto():
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call kccFuncExcel.RangeMergeByValue(Selection, True, True, False, 0)
End Sub

Sub MergePrint(): MsgBox "Print": End Sub
'--------------------------------------------------

'監視開始
'Workbook_Openから呼ばれる
'他ブックの上書き保存を検知するために使用される
Sub MonitorStart(): Set instMergeBlocker = New MergeBlocker: End Sub

'監視停止
Sub MonitorStop(): Set instMergeBlocker = Nothing: End Sub

'セル結合支援自動化ON
Sub Start_MergeCellCreater()
    Call SubMergeCellCreater(False)
End Sub

'セル結合支援自動化OFF
Sub Stop_MergeCellCreater()
    Call SubMergeCellCreater(True)
End Sub

'結合セル一覧表示後に呼び出すプロシージャ
'他ブックのセル選択を検知しするために使用される
Sub CellHighlightStart(): Call CellHighlightStartWs(ActiveSheet): End Sub
Sub CellHighlightStartWs(Optional ws As Worksheet)
    Set instCellHighlighter = New CellHighlighter
    instCellHighlighter.Init ws
End Sub
