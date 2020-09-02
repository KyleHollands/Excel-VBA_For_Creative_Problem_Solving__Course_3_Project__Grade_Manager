Attribute VB_Name = "Module5"
Option Explicit
Option Base 1

Sub ReformatGradeManager()

Dim tWb As Workbook

Sheets(1).Select: Rows("1:1").Select
With Selection.Font
    .Name = "Calibri"
    .Size = 13
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
    .Bold = True
End With

Sheets(1).Select: Range("B1:Z10000").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Cells.Select
Cells.EntireColumn.AutoFit

End Sub
