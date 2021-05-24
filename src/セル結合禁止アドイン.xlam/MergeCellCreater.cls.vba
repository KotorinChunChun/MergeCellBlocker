VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergeCellCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        MergeCellCreater
Rem
Rem  @description   MergeCellCreater
Rem
Rem  @update        2021/05/24
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Rem オブジェクトの作成
Public Function Init(pApp As Excel.Application) As MergeCellCreater
    If Me Is MergeCellCreater Then
        With New MergeCellCreater
            Set Init = .Init(pApp)
        End With
        Exit Function
    End If
    Set Init = Me
    Set app = pApp
End Function

Private Sub app_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Target.MergeCells Then
        Cancel = True
        Call kccFuncExcel.RangeUnMerge(Target)
    End If
End Sub

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call kccFuncExcel.RangeMergeByValue(Target, True, True, True, 0)
End Sub
