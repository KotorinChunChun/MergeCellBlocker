Attribute VB_Name = "CustomUI"
Rem
Rem CustomUI
Rem
Rem 本モジュールは自作のCustomUIエディタから自動生成したイベントハンドラです。
Rem

Sub onAction_AddinStart(control As IRibbonControl): Call AddinStart: FinalUseCommand = "AddinStart": End Sub
Sub onAction_AddinStop(control As IRibbonControl): Call AddinStop: FinalUseCommand = "AddinStop": End Sub
Sub onAction_MergeSearch(control As IRibbonControl): Call MergeSearch: FinalUseCommand = "MergeSearch": End Sub
Sub onAction_MergeBreak(control As IRibbonControl): Call MergeBreak: FinalUseCommand = "MergeBreak": End Sub
Sub onAction_MergeDown(control As IRibbonControl): Call MergeDown: FinalUseCommand = "MergeDown": End Sub
Sub onAction_MergeRight(control As IRibbonControl): Call MergeRight: FinalUseCommand = "MergeRight": End Sub
Sub onAction_MergeAuto(control As IRibbonControl): Call MergeAuto: FinalUseCommand = "MergeAuto": End Sub
Sub onAction_MergePrint(control As IRibbonControl): Call MergePrint: FinalUseCommand = "MergePrint": End Sub
Sub onAction_AddinConfig(control As IRibbonControl): Call AddinConfig: FinalUseCommand = "AddinConfig": End Sub
Sub onAction_AddinInfo(control As IRibbonControl): Call AddinInfo: FinalUseCommand = "AddinInfo": End Sub
Sub onAction_AddinEnd(control As IRibbonControl): Call AddinEnd: FinalUseCommand = "AddinEnd": End Sub



