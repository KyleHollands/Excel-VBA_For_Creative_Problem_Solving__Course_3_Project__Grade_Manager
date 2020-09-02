Attribute VB_Name = "Module2"
Option Base 1
Option Explicit

Sub EventMacro()

Dim path As String, TodayDate As String, strFolderName As String, strFolderExists As String
Dim NowArray As Variant
Dim alertTime As Double
Dim i As Integer
Dim workbookName As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error GoTo Here

path = ActiveWorkbook.path & "\" & "Backups"
NowArray = Split(Now(), " ")
TodayDate = Replace(NowArray(0), "/", "-")
    
i = i + 1

strFolderName = path
strFolderExists = Dir(strFolderName, vbDirectory)
    
If Dir(path & "\", vbDirectory) = "" Then
workbookName = ActiveWorkbook.Name

Else:

    With ThisWorkbook
        .SaveCopyAs path & "\Grade Manager_" & TodayDate & " (" & i & ")" & ".xlsm"
    End With
    
End If

alertTime = Now + TimeValue("00:01:00")
Application.OnTime alertTime, "EventMacro"

Here:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
