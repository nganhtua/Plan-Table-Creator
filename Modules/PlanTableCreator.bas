Attribute VB_Name = "PlanTableCreator"
Option Explicit

Public Const PLANTBL_PAD1 = 3
Public Const PLANTBL_PAD2 = 10
Public Const DEM = vbLf
Public Const MOY = 12

Public Sub PlanTableCreator()
    Dim equip_info, plan_data, plan_table As ListObject
    Dim new_row As ListRow
    Dim plan_range As Range
    Dim r, matched_r, month As Integer
    Dim equip_code, action, current_action, updated_action, _
        equip_no, no_lookup_fml As String
    Dim StartTime As Double
    
    Set equip_info = Sheets("EquipmentInfo").ListObjects("EquipmentInfo")
    Set plan_data = Sheets("PlanData").ListObjects("PlanData")
    Set plan_range = plan_data.DataBodyRange
    Set plan_table = Sheets("PlanTable").ListObjects("PlanTable")
    
    StartTime = Timer
    Range("K2").Value = "Running..."
    Range("K2").Interior.ColorIndex = 6     'Yellow
    Range("K3").ClearContents
    
    With plan_table
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
        .ListRows.Add
        .DataBodyRange(1, PLANTBL_PAD1 + 1).Value = plan_range(1, 2).Value
    End With
    
    For r = 1 To plan_data.ListRows.Count
        equip_code = plan_data.DataBodyRange(r, 2)
        month = plan_data.DataBodyRange(r, 3)
        action = plan_data.DataBodyRange(r, 4)
        matched_r = _
            Application.Match(equip_code, plan_table.ListColumns(PLANTBL_PAD1 + 1).DataBodyRange, 0)
        If IsError(matched_r) Then
            'Chua co thiet bi
            Set new_row = plan_table.ListRows.Add
            new_row.Range(PLANTBL_PAD1 + 1) = equip_code
            matched_r = plan_table.ListRows.Count
        End If
        current_action = _
            plan_table.ListColumns(month + PLANTBL_PAD2).DataBodyRange(matched_r, 1).Value
        If Len(current_action) = 0 Then
            updated_action = action
        Else
            If InStr(1, current_action, action) = 0 Then
                updated_action = current_action & DEM & action
            End If
        End If
        plan_table.ListColumns(month + PLANTBL_PAD2).DataBodyRange(matched_r, 1).Value = _
            updated_action
    Next r
    
    'Getting 'Ma thiet bi' text
    equip_no = equip_info.HeaderRowRange(2).Value
    
    'Set formula for 'TT' column
    no_lookup_fml = "=MATCH([@[" & equip_no & "]],EquipmentInfo[" & equip_no & "],0)"
    plan_table.DataBodyRange(2).Formula = no_lookup_fml
    
    Call OptimizeTable(plan_table.ListColumns(PLANTBL_PAD2 + 1).DataBodyRange.Resize(, MOY))
        
    Range("K2").Value = "Done!"
    Range("K2").Interior.ColorIndex = 4     'Green
    Range("K3").Value = Round(Timer - StartTime, 2)
End Sub

Private Sub OptimizeTable(tbl As Range)
    Dim i, c_index As Integer
    Dim c As Range
    Dim s As String
    Dim after_arr() As String
    
    ReDim after_arr(tbl.Count)
    
    i = 0
    For Each c In tbl
        after_arr(i) = SortPlan(c.Value)
        i = i + 1
    Next c
    
    For i = 1 To tbl.Count
        'Debug.Print tbl.Cells(i \ 12 + 1, i Mod 12).Address
        'Debug.Print after_arr(i - 1)
        If i Mod MOY = 0 Then
            tbl.Cells(i \ MOY, MOY).Value = after_arr(i - 1)
        Else
            tbl.Cells(i \ MOY + 1, i Mod MOY).Value = after_arr(i - 1)
        End If
    Next i
    
    'For i = PLANTBL_PAD2 + 1 To PLANTBL_PAD2 + 13
    '    For Each c In tbl.ListColumns(i).DataBodyRange
    '        c.Value = SortPlan(c.Value)
    '    Next c
    'Next i
End Sub

Private Function SortPlan(month_plan As String) As String
    Dim std_str As String
    Dim std_arr() As String
    Dim before_arr() As String
    Dim after_arr() As String
    Dim i, j As Integer
    
    std_str = "K" & ChrW(272) & DEM & "K" & ChrW(272) & " cân" & DEM & _
        "K" & ChrW(272) & " an toàn" & DEM & "K" & ChrW(272) & " áp k" & _
        ChrW(7871) & DEM & "HC" & DEM & "BT" & DEM & "PQ"
    std_arr = Split(std_str, DEM)
    before_arr = Split(month_plan, DEM)

    j = 0
    If UBound(before_arr) = LBound(before_arr) Then
        SortPlan = month_plan
    ElseIf UBound(before_arr) - LBound(before_arr) > 0 Then
        For i = LBound(std_arr, 1) To UBound(std_arr, 1)
            'If InStr(1, month_plan, std_arr(i)) <> 0 Then
            '    ReDim Preserve after_arr(0 To j) As String
            '    after_arr(j) = std_arr(i)
            '    j = j + 1
            'End If
            If Not IsError(Application.Match(std_arr(i), before_arr, 0)) Then
                ReDim Preserve after_arr(0 To j) As String
                after_arr(j) = std_arr(i)
                j = j + 1
            End If
        Next i
        SortPlan = Join(after_arr, DEM)
    End If
End Function

Private Sub TestSub()
    Dim arr As Variant
    Dim i As Integer
    arr = Array(1)
    Debug.Print Range("J6:U8").Count
    For i = 1 To Range("J6:U8").Count
        Debug.Print i
    Next i
    'Call OptimizeTable(Range("J6:J8").Resize(, MOY))
End Sub
