VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Add Assignment"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3 - Add Assignment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim folder As String, wb As Workbook, FileName As String, CategoryName As String
Dim FindAddress As String, ColLetter As String
Dim aWB As Workbook, i As Integer, nItems As Integer, nInCategory As Integer
Dim FindArray As Variant
Dim myStr As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

CategoryName = UserForm3.homeworkType.Text

'SECTION SHEETS--------------------------------------------------------------------


folder = ActiveWorkbook.path & "\" & "Section Files"

FileName = Dir(folder & "\*.xlsx")

Do
    Workbooks.Open folder & "\" & FileName
    Set aWB = ActiveWorkbook
    
    nItems = WorksheetFunction.CountA(Rows(1))
    Range("B1").Select
    nInCategory = 0
    
    For i = 1 To nItems
        If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = CategoryName Then
            nInCategory = nInCategory + 1
            
            myStr = onlyDigits(ActiveCell.Offset(0, i - 1))
            
            If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = CategoryName _
                And ((Left(ActiveCell.Offset(0, i), Len(CategoryName)) <> CategoryName) Or IsEmpty(ActiveCell.Offset(0, i))) Then
                FindAddress = ActiveCell.Offset(0, i).Address
                FindArray = Split(FindAddress, "$")
                ColLetter = FindArray(1)
                Columns(ColLetter & ":" & ColLetter).Select
                Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                Range(ColLetter & "1") = CategoryName & " " & myStr + 1
                Exit For
            End If
        
        ElseIf Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) <> CategoryName Then
            
            If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = "" Then
                FindAddress = ActiveCell.Offset(0, i - 1).Address
                FindArray = Split(FindAddress, "$")
                ColLetter = FindArray(1)
                Columns(ColLetter & ":" & ColLetter).Select
                Range(ColLetter & "1") = CategoryName & " " & nInCategory + 1
            End If
            
        End If
        
    Next i
    
    Range("A1").Select
    
    aWB.Close SaveChanges:=True
    FileName = Dir
    
Loop Until FileName = ""

'GRADE MANAGER--------------------------------------------------------------------

nItems = WorksheetFunction.CountA(Rows(1))
Range("B1").Select
nInCategory = 0
    
For i = 1 To nItems
    If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = CategoryName Then
        nInCategory = nInCategory + 1
        
        myStr = onlyDigits(ActiveCell.Offset(0, i - 1))
        
        If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = CategoryName _
            And ((Left(ActiveCell.Offset(0, i), Len(CategoryName)) <> CategoryName) Or IsEmpty(ActiveCell.Offset(0, i))) Then
            FindAddress = ActiveCell.Offset(0, i).Address
            FindArray = Split(FindAddress, "$")
            ColLetter = FindArray(1)
            Columns(ColLetter & ":" & ColLetter).Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range(ColLetter & "1") = CategoryName & " " & myStr + 1
            Exit For
        End If
        
    ElseIf Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) <> CategoryName Then
            
        If Left(ActiveCell.Offset(0, i - 1), Len(CategoryName)) = "" Then
            FindAddress = ActiveCell.Offset(0, i - 1).Address
            FindArray = Split(FindAddress, "$")
            ColLetter = FindArray(1)
            Columns(ColLetter & ":" & ColLetter).Select
            Range(ColLetter & "1") = CategoryName & " " & nInCategory + 1
        End If
            
    End If
        
Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "A column for " & CategoryName & nInCategory + 1 & " has been added to the files!"

End Sub

Private Sub CommandButton2_Click()

Unload UserForm3

End Sub

Function onlyDigits(s As String) As String

Dim retval As String
Dim i As Integer

retval = ""

For i = 1 To Len(s)
    If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
        retval = retval + Mid(s, i, 1)
    End If
Next
                          '
onlyDigits = retval

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then

    Cancel = True
    'Do Nothing

End If

End Sub



