Attribute VB_Name = "CustomUI"
Rem
Rem CustomUI
Rem
Rem 本モジュールは自作のCustomUIエディタから自動生成したイベントハンドラです。
Rem

Sub onAction_Start_MergeCellBlocker(Control As IRibbonControl): Call Start_MergeCellBlocker: FinalUseCommand = "Start_MergeCellBlocker": End Sub
Sub onAction_Stop_MergeCellBlocker(Control As IRibbonControl): Call Stop_MergeCellBlocker: FinalUseCommand = "Stop_MergeCellBlocker": End Sub
Sub onAction_MergeSearch(Control As IRibbonControl): Call MergeSearch: FinalUseCommand = "MergeSearch": End Sub

Sub onAction_Start_MergeCellCreater(Control As IRibbonControl): Call Start_MergeCellCreater: FinalUseCommand = "Start_MergeCellCreater": End Sub
Sub onAction_Stop_MergeCellCreater(Control As IRibbonControl): Call Stop_MergeCellCreater: FinalUseCommand = "Stop_MergeCellCreater": End Sub
Sub onAction_MergeAuto(Control As IRibbonControl): Call MergeAuto: FinalUseCommand = "MergeAuto": End Sub
Sub onAction_MergeDown(Control As IRibbonControl): Call MergeDown: FinalUseCommand = "MergeDown": End Sub
Sub onAction_MergeRight(Control As IRibbonControl): Call MergeRight: FinalUseCommand = "MergeRight": End Sub
Sub onAction_MergeBreak(Control As IRibbonControl): Call MergeBreak: FinalUseCommand = "MergeBreak": End Sub

Sub onAction_MergePrint(Control As IRibbonControl): Call MergePrint: FinalUseCommand = "MergePrint": End Sub
Sub onAction_AddinConfig(Control As IRibbonControl): Call AddinConfig: FinalUseCommand = "AddinConfig": End Sub
Sub onAction_AddinInfo(Control As IRibbonControl): Call AddinInfo: FinalUseCommand = "AddinInfo": End Sub
Sub onAction_AddinEnd(Control As IRibbonControl): Call AddinEnd: FinalUseCommand = "AddinEnd": End Sub
