Attribute VB_Name = "Module4"
Option Explicit
Option Base 1

Sub GradeManager()

Dim tWb As Workbook
Dim sheetFound As Boolean
Dim sheetToFind As String

Set tWb = ThisWorkbook

sheetToFind = "Roster"
sheetFound = sheetExists(sheetToFind)

If sheetFound = False Then
    
    UserForm2.Height = 131.25
    UserForm2.Show
    
ElseIf sheetFound = True And Sheets("Roster").Range("A1") = "" Then

    UserForm2.Height = 131.25
    UserForm2.Show
    
ElseIf sheetFound = True Then

    UserForm2.Height = 224
    UserForm2.CommandButton7.Visible = False
    UserForm2.CommandButton8.Visible = False
    
    UserForm2.CommandButton1.Top = 50
    UserForm2.CommandButton2.Top = 50
    UserForm2.CommandButton3.Top = 98
    UserForm2.CommandButton4.Top = 98
    UserForm2.CommandButton5.Top = 146
    UserForm2.CommandButton6.Top = 146
    
    UserForm2.Label1.Top = 14
    
    UserForm2.Show

End If

End Sub

Function sheetExists(sheetToFind As String) As Boolean

Dim Sheet As Worksheet

sheetExists = False

For Each Sheet In Worksheets
    If sheetToFind = Sheet.Name Then
        sheetExists = True
        Exit Function
    End If
Next Sheet
    
End Function
