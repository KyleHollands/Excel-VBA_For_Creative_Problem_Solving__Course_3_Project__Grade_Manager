Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public folder As String, NewFolderName As String

Sub CourseSetup()

Dim tWb As Workbook, aWB As Workbook
Dim RosterFileName As String, sName As String, NewFileName As String
Dim ws As Worksheet
Dim sheetExists As Boolean, Ans As Integer
Dim minSection As Integer, maxSection As Integer, nNames As Integer, nSections As Integer
Dim i As Integer, j As Integer, k As Integer, allSections As Integer
Dim N() As String, FileName As String, sectFolder As String
Dim sectNum As Variant, h As Variant
Dim SI() As Double

On Error GoTo Here

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set tWb = ThisWorkbook
RosterFileName = tWb.path & "\" & "Roster" & ".xlsx"

'CREATING THE COURSE FOLDER------------------------------------------------------

NewFolderName = InputBox("Please enter a name for the new folder: ")

If (StrPtr(NewFolderName) = 0) Then
    GoTo Here

ElseIf (NewFolderName = "") Then
    MsgBox "You did not enter anything"
    GoTo Here
Else
    'Do Nothing
End If


MsgBox ("Please choose the directory/folder where you'd like to place the course folder." _
    & vbCrLf & "A new folder named " & "'" & NewFolderName & "'" & " will be created in this directory.")

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Show
    folder = .SelectedItems(1)
End With
    
MkDir folder & "\" & NewFolderName

MsgBox ("New folder " & "'" & NewFolderName & "'" & " created in " & folder & ".")

NewFileName = NewFolderName

tWb.SaveAs FileName:= _
    folder & "\" & NewFolderName & "\" & NewFileName & ".xlsm" _
    , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    
MkDir (folder & "\" & NewFolderName & "\" & "Section Files")
MkDir (folder & "\" & NewFolderName & "\" & "Backups")

'IMPORTING THE ROSTER FILE FOR THE COURSE------------------------------------------------------

Sheets.Add After:=ActiveSheet: ActiveSheet.Select: ActiveSheet.Name = "Roster"

MsgBox "Navigate to the intitial class roster. This file must have student names in column A" _
& "(Last Name, First Name), student numbers in column B, and section numbers in column C."

RosterFileName = Application.GetOpenFilename(FileFilter:="Excel Filter (.xlsx),*.xlsx", Title:="Open Roster File")
Workbooks.Open RosterFileName

Set aWB = ActiveWorkbook

aWB.Sheets("Sheet1").Range("A:C").Select: Selection.Copy: aWB.Close SaveChanges:=False
tWb.Sheets("Roster").Range("A:C").Select: Worksheets("Roster").Paste

'CREATE THE SECTION FILES------------------------------------------------------

Set tWb = ThisWorkbook

minSection = WorksheetFunction.Min(tWb.Sheets("Roster").Columns("C:C"))
maxSection = WorksheetFunction.Max(tWb.Sheets("Roster").Columns("C:C"))

nNames = WorksheetFunction.CountA(tWb.Sheets("Roster").Columns("A:A"))
nSections = maxSection - minSection + 1

For i = 1 To nSections
    k = 0
    minSection = (minSection + 1)
    
    ReDim N(1) As String
    ReDim SI(1) As Double
    
    For j = 1 To nNames
        If tWb.Sheets("Roster").Range("C" & j) = (minSection - 1) Then
            sectNum = tWb.Sheets("Roster").Range("C" & j)
            k = k + 1
            ReDim Preserve N(k)
            ReDim Preserve SI(k)
            N(k) = tWb.Sheets("Roster").Range("A" & j)
            SI(k) = tWb.Sheets("Roster").Range("B" & j)
        End If
    Next j
    
    If sectNum = (minSection - 1) Then
    
        Workbooks.Add
        
        Set aWB = ActiveWorkbook
        
        Range("A1") = "Name"
        Range("B1") = "Student ID"
        
        For j = 1 To k
            Range("A" & j).Offset(1, 0) = N(j)
            Range("A" & j).Offset(1, 1) = SI(j)
        Next j
        
        ActiveWorkbook.SaveAs FileName:= _
            folder & "\" & NewFolderName & "\" & "Section Files" & "\Section_" & sectNum & ".xlsx" _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            
        aWB.Close SaveChanges:=False
    
    End If
    
Next i

Call PopulateComboBoxes

UserForm1.Show

Here:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
