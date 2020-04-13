Attribute VB_Name = "FuncString_Partial"
'����񂿂�񃉃C�u������FuncString����؂�o�����@�\
Option Explicit

Rem �Z���Q�Ǝ����u�b�N�A�V�[�g�A�Z���ɕ�������
Rem
Rem  @param base_adr_string �t�H�[�~�����Q�ƕ�����
Rem
Rem  @return String(0 To 3) {�p�X,�u�b�N,�V�[�g,�����W}��Ԃ��B
Rem
Rem  @example
Rem         Missing                                             >> String(0 to 3) {"","","",""}
Rem         String ""                                           >> String(0 to 3) {"","","",""}
Rem         String "D22"                                        >> String(0 to 3) {"","","","D22"}
Rem         String "=D22"                                       >> String(0 to 3) {"","","","D22"}
Rem         String "=Sheet2!C17"                                >> String(0 to 3) {"","","Sheet2","C17"}
Rem         String "='She!et'!C24"                              >> String(0 to 3) {"","","She!et","C24"}
Rem         String "=[�ق��ق�.xls]�ꗗ!$E$27"                  >> String(0 to 3) {"","�ق��ق�.xls","�ꗗ","$E$27"}
Rem         String "='[�ق��ق�.xls]She!et'!$H$36"              >> String(0 to 3) {"","�ق��ق�.xls","She!et","$H$36"}
Rem         String "='C:\[���[�N�V�[�g.xlsx]Sheet1'!$C$6"       >> String(0 to 3) {"C:\","���[�N�V�[�g.xlsx","Sheet1","$C$6"}
Rem         String "='C:\[���[�N�V�[�g.xlsx]She!et'!$B$7"       >> String(0 to 3) {"C:\","���[�N�V�[�g.xlsx","She!et","$B$7"}
Rem         String "='[Book(a)1.xlsx]Sheet1'!$C$15"             >> String(0 to 3) {"","Book(a)1.xlsx","Sheet1","$C$15"}
Rem         String "='C:\Folder\[Book[a]1.xlsx]Sheet1'!$C$15"   >> String(0 to 3) {"C:\Folder\","Book[a]1.xlsx","Sheet1","$C$15"}
Rem         String "='C:\Folder\[[Booka]1.xlsx]Sheet1'!$C$15"   >> String(0 to 3) {"C:\Folder\","[Booka]1.xlsx","Sheet1","$C$15"}
Rem         String "='[Book''a''2.xlsx]Sheet1'!$B$11"           >> String(0 to 3) {"","Book'a'2.xlsx","Sheet1","$B$11"}
Rem         String "='C:\Folder\[Book''a''2.xlsx]Sheet1'!$B$11" >> String(0 to 3) {"C:\Folder\","Book'a'2.xlsx","Sheet1","$B$11"}
Rem         String "='C:\Folder[a]1\[Book2.xlsx]Sheet1'!$B$18"  >> String(0 to 3) {"C:\Folder[a]1\","Book2.xlsx","Sheet1","$B$18"}
Rem         String "=[Book3.xlsx]Sheet1!$B$14"                  >> String(0 to 3) {"","Book3.xlsx","Sheet1","$B$14"}
Rem
Rem  @note
Rem    �p�X�������u�b�N�� [Book[a]1.xlsx] �́AFormula�擾���_�� [Book(a)1.xlsx] �ƕω����邪�l�����Ă��Ȃ��B
Rem    If s(0) = "" Then book = Replace(Replace(s(1), "(", "["), ")", "]")
Rem    �Ƃ��ׂ���������Ȃ����A�ۊ��ʂ�������̂ł��蓾�Ȃ��B�������Ή�
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

    ''[�ق��ق�.xls]She!et' >> [�ق��ق�.xls]She!et
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
        '[�ق��ق�.xls]�ꗗ >> �ق��ق�.xls , �ꗗ
        book = Mid(s, 2, InStrRev(s, "]") - 2)
        sheet = Right(s, Len(s) - 2 - Len(book))
    Else
        book = ""
        sheet = s
    End If

    '�V���O���N�H�[�e�[�V�����΍�
    Path = Replace(Path, "''", "'")
    book = Replace(book, "''", "'")

    ss(0) = Path
    ss(1) = book
    ss(2) = sheet
    ss(3) = cell
    SplitFormulaPathBookSheetCell = ss
End Function

Rem �Z���Q�Ǝ���Range�I�u�W�F�N�g�ɕϊ�����
Function GetRangeByFormula(formula_str) As Excel.Range
    Dim v: v = SplitFormulaPathBookSheetCell(formula_str)
    Set GetRangeByFormula = Workbooks(v(1)).Worksheets(v(2)).Range(v(3))
End Function

