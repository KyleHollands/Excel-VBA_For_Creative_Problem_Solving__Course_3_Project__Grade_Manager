Attribute VB_Name = "Module3"
Option Explicit
Option Base 1

Sub PopulateComboBoxes()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim i As Integer

For i = 1 To 10
    UserForm1.assignComboBox.AddItem i
    UserForm1.examComboBox.AddItem i
    UserForm1.labComboBox.AddItem i
Next i

UserForm1.labComboBox.Text = "1"
UserForm1.assignComboBox.Text = "1"
UserForm1.examComboBox.Text = "1"

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
