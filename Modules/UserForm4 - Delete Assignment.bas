VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm4 - Delete Assignment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()

Dim folder As String, wb As Workbook, FileName As String, CategoryName As String
Dim FindAddress As String, ColLetter As String
Dim aWB As Workbook, i As Integer, nItems As Integer, nInCategory As Integer
Dim FindArray As Variant

Application.ScreenUpdating = False

CategoryName = UserForm4.homeworkType.Text

'SECTION SHEETS--------------------------------------------------------------------


folder = ActiveWorkbook.path & "\" & "Section Files"

FileName = Dir(folder & "\*.xlsx")

Do
    Workbooks.Open folder & "\" & FileName
    Set aWB = ActiveWorkbook
    
    nItems = WorksheetFunction.CountA(Rows(1))
    Range("C1").Select
    nInCategory = 0
    
    For i = 1 To nItems
        If ActiveCell.Offset(0, i - 1) = CategoryName Then
            ActiveCell.Offset(0, i - 1).EntireColumn.Delete
        End If
    Next i
    
    Range("A1").Select
    
    aWB.Close SaveChanges:=True
    FileName = Dir
    
Loop Until FileName = ""

'GRADE MANAGER--------------------------------------------------------------------

nItems = WorksheetFunction.CountA(Rows(1))
    Range("D1").Select
    nInCategory = 0
    
    For i = 1 To nItems
        If ActiveCell.Offset(0, i - 1) = CategoryName Then
            ActiveCell.Offset(0, i - 1).EntireColumn.Delete
        End If
    Next i
            
Application.ScreenUpdating = True

MsgBox "Column " & CategoryName & " has been deleted from the files!"

End Sub

Private Sub CommandButton2_Click()

Unload UserForm4

End Sub

Private Sub CommandButton4_Click()

Unload UserForm4

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then

    Cancel = True
    'Do Nothing

End If

End Sub




