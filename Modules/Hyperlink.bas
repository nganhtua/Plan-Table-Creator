Attribute VB_Name = "Hyperlink"
Option Explicit

Sub InsertHyperlinkToCell()
    'Choose a file and link it to the selected cells
    'Default folder determined in cell I1
    
    Dim hyperlink_fullpath, filename As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .ButtonName = "Select"
        .InitialFileName = Range("J1").Value 'Application.DefaultFilePath
        .Title = "Select File"
        If .Show = 0 Then Exit Sub
        hyperlink_fullpath = .SelectedItems(1)
    End With
    
    Dim c As Range
    Dim fontname As String
    For Each c In Selection
        fontname = c.Font.Name
        ActiveSheet.Hyperlinks.Add Anchor:=c, _
            Address:=hyperlink_fullpath, _
            TextToDisplay:=Mid(hyperlink_fullpath, InStrRev(hyperlink_fullpath, "\") + 1) 'change the text to display as desired
        c.Font.Name = fontname
    Next c
    Range("J1").Value = Left(hyperlink_fullpath, InStrRev(hyperlink_fullpath, "\"))
End Sub

Function CheckHyperlink(cell_address As Range) As Boolean
    'Check the availability of the hyperlink in a cell
    
    Dim objFSO As Object
    Dim lnk As String
    If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
    lnk = cell_address.Hyperlinks(1).Address
    CheckHyperlink = objFSO.FileExists(lnk)
End Function

Sub OpenFileLocation()
    'Open filepath (directory) linked in ActiveCell
    
    Dim fullpath As String
    Dim fso As Object
    Dim fname, fpath As String
    
    On Error Resume Next
    fullpath = ActiveCell.Hyperlinks(1).Address
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fullpath) Or fso.FolderExists(fullpath) Then
        fpath = Left(fullpath, InStrRev(fullpath, "\"))
        fname = Mid(fullpath, InStrRev(fullpath, "\") + 1)
        'Call Shell("explorer.exe" & " " & fpath, vbNormalFocus)
        ThisWorkbook.FollowHyperlink fpath
    Else
        MsgBox "This file doesn't exist anymore. Please make sure the path is correct.", vbExclamation
    End If
End Sub

Sub test()
    'Testing module, I can do many crazy things here
    
    Dim c As Range
    Dim filename As String
    For Each c In Range("F9:F31")
        c.Value = Mid(c.Value, InStrRev(c.Value, "\") + 1)
    Next c
End Sub
