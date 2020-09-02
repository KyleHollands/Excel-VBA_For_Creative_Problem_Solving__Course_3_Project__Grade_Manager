VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "UserForm5 - Replace or Add.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Range("D2").Select

If IsEmpty(ActiveCell.Offset(UserForm5.ComboBox1.ListIndex, UserForm5.ComboBox2.ListIndex)) Then
    MsgBox ("No grade found.")
Else:
    MsgBox ("Grade on this assignment is " & ActiveCell.Offset(UserForm5.ComboBox1.ListIndex, UserForm5.ComboBox2.ListIndex))
End If

End Sub

Private Sub CommandButton2_Click()

Dim NewGrade As Double, Ans As Integer
Dim temp As String
Dim tempRange As Range

Application.ScreenUpdating = False
Application.DisplayAlerts = False

NewGrade = InputBox("Please enter the new grade: ")

Range("D2").Select

Ans = MsgBox("Are you sure you want to replace/add this grade?", vbYesNo)

If Ans = 6 Then
    ActiveCell.Offset(UserForm5.ComboBox1.ListIndex, UserForm5.ComboBox2.ListIndex) = NewGrade
'    MsgBox ("Grade replaced/added to: " & NewGrade)
End If

path = ActiveWorkbook.path & "\" & "Section Files"
FileName = Dir(path & "\" & "*.xlsx")

Do
    Workbooks.Open path & "\" & FileName
    Set aWB = ActiveWorkbook
    nStudents = WorksheetFunction.CountA(aWB.Sheets(1).Columns("A:A")) - 1: Range("A2").Select
    
    Range("C2").Select
    
    If Ans = 6 Then
        temp = UserForm5.ComboBox1.Value
'        If Not IsError(Application.Match(temp, Sheets("Sheet1").Range("A:A"), 0)) Then
            For Each t In Range("A:A")
                If t = temp Then
                    t.Offset(0, UserForm5.ComboBox2.ListIndex + 2) = NewGrade
                    Exit For
                End If
                If t = "" Then
                    Exit For
                End If
            Next t
    End If
    
    aWB.Close SaveChanges:=True
    
    FileName = Dir
        
Loop Until FileName = ""

MsgBox (temp & " grade" & " for " & UserForm5.ComboBox2.Text & " added/modified to: " & NewGrade)

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton3_Click()

Unload UserForm5

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then

    Cancel = True
    'Do Nothing
End If

End Sub


