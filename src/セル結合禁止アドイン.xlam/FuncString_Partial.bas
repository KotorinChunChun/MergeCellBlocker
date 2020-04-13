Attribute VB_Name = "FuncString_Partial"
'ちゅんちゅんライブラリのFuncStringから切り出した機能
Option Explicit

Rem セル参照式をブック、シート、セルに分割する
Rem
Rem  @param base_adr_string フォーミュラ参照文字列
Rem
Rem  @return String(0 To 3) {パス,ブック,シート,レンジ}を返す。
Rem
Rem  @example
Rem         Missing                                             >> String(0 to 3) {"","","",""}
Rem         String ""                                           >> String(0 to 3) {"","","",""}
Rem         String "D22"                                        >> String(0 to 3) {"","","","D22"}
Rem         String "=D22"                                       >> String(0 to 3) {"","","","D22"}
Rem         String "=Sheet2!C17"                                >> String(0 to 3) {"","","Sheet2","C17"}
Rem         String "='She!et'!C24"                              >> String(0 to 3) {"","","She!et","C24"}
Rem         String "=[ほげほげ.xls]一覧!$E$27"                  >> String(0 to 3) {"","ほげほげ.xls","一覧","$E$27"}
Rem         String "='[ほげほげ.xls]She!et'!$H$36"              >> String(0 to 3) {"","ほげほげ.xls","She!et","$H$36"}
Rem         String "='C:\[ワークシート.xlsx]Sheet1'!$C$6"       >> String(0 to 3) {"C:\","ワークシート.xlsx","Sheet1","$C$6"}
Rem         String "='C:\[ワークシート.xlsx]She!et'!$B$7"       >> String(0 to 3) {"C:\","ワークシート.xlsx","She!et","$B$7"}
Rem         String "='[Book(a)1.xlsx]Sheet1'!$C$15"             >> String(0 to 3) {"","Book(a)1.xlsx","Sheet1","$C$15"}
Rem         String "='C:\Folder\[Book[a]1.xlsx]Sheet1'!$C$15"   >> String(0 to 3) {"C:\Folder\","Book[a]1.xlsx","Sheet1","$C$15"}
Rem         String "='C:\Folder\[[Booka]1.xlsx]Sheet1'!$C$15"   >> String(0 to 3) {"C:\Folder\","[Booka]1.xlsx","Sheet1","$C$15"}
Rem         String "='[Book''a''2.xlsx]Sheet1'!$B$11"           >> String(0 to 3) {"","Book'a'2.xlsx","Sheet1","$B$11"}
Rem         String "='C:\Folder\[Book''a''2.xlsx]Sheet1'!$B$11" >> String(0 to 3) {"C:\Folder\","Book'a'2.xlsx","Sheet1","$B$11"}
Rem         String "='C:\Folder[a]1\[Book2.xlsx]Sheet1'!$B$18"  >> String(0 to 3) {"C:\Folder[a]1\","Book2.xlsx","Sheet1","$B$18"}
Rem         String "=[Book3.xlsx]Sheet1!$B$14"                  >> String(0 to 3) {"","Book3.xlsx","Sheet1","$B$14"}
Rem
Rem  @note
Rem    パスが無いブック名 [Book[a]1.xlsx] は、Formula取得時点で [Book(a)1.xlsx] と変化するが考慮していない。
Rem    If s(0) = "" Then book = Replace(Replace(s(1), "(", "["), ")", "]")
Rem    とすべきかもしれないが、丸括弧が消えるのであり得ない。事実上非対応
Rem
Public Function SplitFormulaPathBookSheetCell(ByVal base_str As Variant) As Variant
    Dim ss(0 To 3) As String

    SplitFormulaPathBookSheetCell = ss
    If IsMissing(base_str) Then Exit Function
    If base_str = "" Then Exit Function

    Dim Path, book, sheet, cell
    Dim s: s = base_str
    If s Like "=*" Then s = Mid(s, 2, Len(s) - 1)

    If s Like "*!*" Then
        cell = Mid(s, InStrRev(s, "!") + 1)
        s = Left(s, Len(s) - Len(cell) - 1)
    Else
        cell = s
        s = ""
    End If

    ''[ほげほげ.xls]She!et' >> [ほげほげ.xls]She!et
    If s Like "'*'" Then
        s = Mid(s, 2, Len(s) - 2)
    End If

    'C:\Folder\[[Booka]1.xlsx]Sheet1 >> [[Booka]1.xlsx]Sheet1
    If s Like "*\*" Then
        Path = Left(s, InStrRev(s, "\"))
        s = Mid(s, Len(Path) + 1)
    Else
        Path = ""
    End If

    If s Like "[[]*[]]*" Then
        '[ほげほげ.xls]一覧 >> ほげほげ.xls , 一覧
        book = Mid(s, 2, InStrRev(s, "]") - 2)
        sheet = Right(s, Len(s) - 2 - Len(book))
    Else
        book = ""
        sheet = s
    End If

    'シングルクォーテーション対策
    Path = Replace(Path, "''", "'")
    book = Replace(book, "''", "'")

    ss(0) = Path
    ss(1) = book
    ss(2) = sheet
    ss(3) = cell
    SplitFormulaPathBookSheetCell = ss
End Function

Rem セル参照式をRangeオブジェクトに変換する
Function GetRangeByFormula(formula_str) As Excel.Range
    Dim v: v = SplitFormulaPathBookSheetCell(formula_str)
    Set GetRangeByFormula = Workbooks(v(1)).Worksheets(v(2)).Range(v(3))
End Function

