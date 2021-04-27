# Macro
MacrosVB

Sub Draft1()
'
' Draft1 Macro
'

'
Range("A1:GG1660").Select
Range("B1").Activate
Selection.Copy
Workbooks.Add

ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Rows("10:10").Select
Selection.AutoFilter

Columns("A:A").Select
Selection.EntireColumn.Hidden = True

Sheets("Sheet1").Select
Sheets("Sheet1").Copy After:=Sheets(1)
Sheets("Sheet1").Select


