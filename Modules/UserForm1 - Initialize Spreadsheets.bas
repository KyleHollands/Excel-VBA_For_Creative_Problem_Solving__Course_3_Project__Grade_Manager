VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Course Setup"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "UserForm1 - Initialize Spreadsheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub CommandButton1_Click()

Dim aWB As Workbook, tWb As Workbook
Dim minSection As Integer, maxSection As Integer, nNames As Integer, nSections As Integer
Dim i As Integer, j As Integer, k As Integer, allSections As Integer
Dim N() As String, FileName As String
Dim sectNum As Variant, h As Variant
Dim SI() As Double

Application.DisplayAlerts = False
Application.ScreenUpdating = False

On Error GoTo Here

'INITIALIZE SECTION FILE INFORMATION------------------------------------------------------

Set tWb = ThisWorkbook

FileName = Dir(folder & "\" & NewFolderName & "\" & "Section Files" & "\*.xlsx")

Do
    Workbooks.Open folder & "\" & NewFolderName & "\" & "Section Files" & "\" & FileName
    Set aWB = ActiveWorkbook
    
    For Each h In Range("A1").EntireRow.Cells
        If IsEmpty(h) Then
            For i = 1 To UserForm1.assignComboBox.ListIndex + 1
                h.Offset(0, i - 1) = "Assignment " & i
            Next i
            Exit For
        End If
    Next h
    
    For Each h In Range("A1").EntireRow.Cells
        If IsEmpty(h) Then
            For i = 1 To UserForm1.examComboBox.ListIndex + 1
                h.Offset(0, i - 1) = "Exam " & i
            Next i
            Exit For
        End If
    Next h
    
    For Each h In Range("A1").EntireRow.Cells
        If IsEmpty(h) Then
            For i = 1 To UserForm1.labComboBox.ListIndex + 1
                h.Offset(0, i - 1) = "Lab " & i
            Next i
            Exit For
        End If
    Next h
    
    Call ReformatGradeManager
    
    aWB.Close SaveChanges:=True
    FileName = Dir
    
Loop Until FileName = ""

'INITIALIZE ROSTER FILE INFORMATION------------------------------------------------------

For Each h In Range("A1").EntireRow.Cells
    If IsEmpty(h) Then
        For i = 1 To UserForm1.assignComboBox.ListIndex + 1
            h.Offset(0, i - 1) = "Assignment " & i
        Next i
        Exit For
    End If
Next h

For Each h In Range("A1").EntireRow.Cells
    If IsEmpty(h) Then
        For i = 1 To UserForm1.examComboBox.ListIndex + 1
            h.Offset(0, i - 1) = "Exam " & i
        Next i
        Exit For
    End If
Next h

For Each h In Range("A1").EntireRow.Cells
    If IsEmpty(h) Then
        For i = 1 To UserForm1.labComboBox.ListIndex + 1
            h.Offset(0, i - 1) = "Lab " & i
        Next i
        Exit For
    End If
Next h

MsgBox ("Assignment headings have been added to the files. Setup is complete.")

Sheets("Grade Manager").Delete

Call ReformatGradeManager

Unload UserForm1
Unload UserForm2

Here:

Range("Z100").Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True

Call GradeManager

End Sub

Private Sub CommandButton2_Click()

Call ReformatGradeManager

Unload UserForm1
Unload UserForm2

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then

    Cancel = True
    'Do Nothing

End If

End Sub
