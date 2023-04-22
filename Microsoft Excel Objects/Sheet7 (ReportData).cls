VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim ATTACHMENT_COL As Byte

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ATTACHMENT_COL = 7
    On Error Resume Next
    ActiveSheet.Buttons("btn_attach_file").Delete
    ActiveSheet.Buttons("btn_open_folder").Delete
    Debug.Print ActiveCell.Hyperlinks(1).Address
    If ActiveCell.Column = ATTACHMENT_COL And Target.Columns.Count = 1 Then
        With ActiveCell.Offset(0, 1)
            ActiveSheet.Buttons.Add(.Left, .Top, .Height * 1.2, .Height).Name = "btn_attach_file"
            ActiveSheet.Buttons("btn_attach_file").OnAction = "InsertHyperlinkToCell"
            ActiveSheet.Buttons("btn_attach_file").Characters.Text = "..."
            ActiveSheet.Buttons("btn_attach_file").TooltipText = "Attach a file"
            ActiveSheet.Buttons.Add(.Left + .Height * 1.2, .Top, .Height * 6.5, .Height).Name = "btn_open_folder"
            ActiveSheet.Buttons("btn_open_folder").OnAction = "OpenFileLocation"
            ActiveSheet.Buttons("btn_open_folder").Characters.Text = "Open file location"
        End With
    End If
End Sub
