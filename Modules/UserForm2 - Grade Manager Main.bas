VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Grade Manager"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
   OleObjectBlob   =   "UserForm2 - Grade Manager Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub CommandButton1_Click()

Dim aWB As Workbook, tWb As Workbook
Dim path As String, Name As String, FileName As String, folder As String, NewFolderName As String
Dim nNames As Integer, nItems As Integer, nStudents As Integer
Dim i As Integer, idx As Integer, j As Integer
Dim G() As Variant

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error GoTo Here

Set tWb = ThisWorkbook

path = ActiveWorkbook.path & "\" & "Section Files"

nNames = WorksheetFunction.CountA(tWb.Sheets(1).Columns("A:A")) - 1
nItems = WorksheetFunction.CountA(tWb.Sheets(1).Rows(1)) - 3

ReDim G(nItems) As Variant
FileName = Dir(path & "\" & "*.xlsx")

Do
    Workbooks.Open path & "\" & FileName
    Set aWB = ActiveWorkbook
    nStudents = WorksheetFunction.CountA(aWB.Sheets(1).Columns("A:A")) - 1: Range("A2").Select
    
    For i = 1 To nStudents
        aWB.Activate
        Name = ActiveCell.Offset(i - 1, 0)
        
        G = Range("C" & i + 1).EntireRow
        ReDim Preserve G(1, nItems + 2) As Variant
        
        tWb.Activate
        idx = WorksheetFunction.Match(Name, tWb.Sheets("Roster").Range("A:A"), 0)
        Sheets("Roster").Range("C" & idx).Select
        
        For j = 1 To nItems
            ActiveCell.Offset(0, j) = G(1, j + 2)
        Next j
        
    Next i
    
    aWB.Close SaveChanges:=True
    
    FileName = Dir
    
Loop Until FileName = ""

MsgBox ("Files synced.")

Here:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton2_Click()

Dim nr As Integer, nc As Integer, i As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error GoTo Here

nr = WorksheetFunction.CountA(Columns("A:A"))
nc = WorksheetFunction.CountA(Rows(1))
Range("A2").Select

For i = 1 To nr - 1
    UserForm5.ComboBox1.AddItem ActiveCell.Offset(i - 1, 0)
Next i

UserForm5.ComboBox1.Text = Range("A2")

Range("D1").Select

For i = 1 To nc - 3
    UserForm5.ComboBox2.AddItem ActiveCell.Offset(0, i - 1)
Next i

UserForm5.ComboBox2.Text = Range("D1")

Application.ScreenUpdating = True
Application.DisplayAlerts = True

Here:

UserForm5.Show

End Sub

Private Sub CommandButton3_Click()

UserForm3.homeworkType.AddItem "Assignment"
UserForm3.homeworkType.AddItem "Exam"
UserForm3.homeworkType.AddItem "Lab"

UserForm3.homeworkType.Text = "Assignment"

UserForm3.Show

End Sub

Private Sub CommandButton4_Click()

Dim i As Variant

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error GoTo Here

Range("C1").Select

For i = 1 To WorksheetFunction.CountA(Rows(1)) - 3
    If i <> "" Then
        UserForm4.homeworkType.AddItem ActiveCell.Offset(0, i)
    Else:
    End If
Next i

UserForm4.homeworkType.Text = Range("D1")

Here:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

UserForm4.Show

End Sub

Private Sub CommandButton5_Click()

Call EventMacro

End Sub

Private Sub CommandButton6_Click()

Unload UserForm2

End Sub

Private Sub CommandButton7_Click()

CourseSetup

End Sub

Private Sub CommandButton8_Click()

Unload UserForm2

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then

    Cancel = True
    'Do Nothing

End If

End Sub

