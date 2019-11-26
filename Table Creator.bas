Attribute VB_Name = "Table Creator"
Public SLICER_ITEM_1 As Variant
Public SLICER_ITEM_2 As Variant

Public Const SLICER_1 As String = ""            ' Should be name of first slicer
Public Const SLICER_2 As String = ""            ' Should be name of second slicer
Public Const ROW As Integer = 22                ' Set ROW to where the first row-cell in your table starts
Public Const COL As Integer = 2                 ' Set COL to where the first column-cell in your table starts
Public Const COPY_CELL As String = "A2"         ' Set to where the value from the pivot you want to copy
Public Const SHEET_NAME As String = "Sheet2"    ' Set to SheetName

Sub initialize_global_variables()
    ' Procedure that initializes the global arrays
    
    ReDim SLICER_ITEM_1(6) ' ReDim to the correct number of items in your first slicer
    ReDim SLICER_ITEM_2(5) ' ReDim to the correct number of items in your second slicer
    
    SLICER_ITEM_1 = Array("", "", "", "", "", "") ' Should be the items in your slicer
    SLICER_ITEM_2 = Array("", "", "", "", "") ' should be the items in another slicer if you have more than one (optional)

End Sub

Sub first_slicer_loop(wb As Workbook, ws1 As Worksheet, slicer_name As String, selected_name As String, array_name As Variant)
    
    '  Procedure that selects each item one by one in the first slicer
    '  wb:  ThisWorkboook
    '  ws:  The sheet where the pivot table is located
    '  slicer_name:  The first/initial slicer being looped through
    '  selected_name:  The selected slicer item (used in the for loop below)
    '  array_name:  The array of the items (SLICER_ITEM_1) being looped through
    
    Dim a_name_1 As Variant
    Dim a_name_2 As Variant
    Dim flag As Boolean
    
    flag = False
    With wb.SlicerCaches(slicer_name)
        For Each a_name_1 In array_name
            If a_name_1 = selected_name Then
                flag = True
                .SlicerItems(a_name_1).Selected = flag
            End If
                flag = False
                For Each a_name_2 In array_name
                    If a_name_2 <> selected_name Then .SlicerItems(a_name_2).Selected = flag
                Next a_name_2
        Next a_name_1
    End With
End Sub

Sub second_slicer_loop(wb As Workbook, ws As Worksheet, slicer_name As String, copy_range As String, _
                       row_index As Integer, col_index As Integer)
    
    '  Procedure that loops through each slicer item in the SLICER_ITEM_2
    '  wb:  ThisWorkboook
    '  ws:  The sheet where the pivot table is located
    '  slicer_name:  The second slicer being looped through

    Dim i As Integer
    Dim sli As slicerItem
    
    With wb.SlicerCaches(slicer_name)
        For Each sli In .VisibleSlicerItems
            If sli.Name <> .SlicerItems(1).Name Then
                sli.Selected = False
            End If
        Next sli
        
        ''''''''''''''''''''''''''''''''''''''''
        '' Put code in this area to move data ''
        ''''''''''''''''''''''''''''''''''''''''
        ws.Cells(row_index, col_index).Value = ws.Range(copy_range).Value
        col_index = col_index + 1
                
        For i = 2 To .SlicerItems.Count
            .SlicerItems(i).Selected = True
            .SlicerItems(i - 1).Selected = False
        ''''''''''''''''''''''''''''''''''''''''
        '' Put code in this area to move data ''
        ''''''''''''''''''''''''''''''''''''''''
            ws.Cells(row_index, col_index).Value = ws.Range(copy_range).Value
            col_index = col_index + 1
        Next i
        ''''''''''''''''''''''''''''''''''''''''
        '' Put code in this area to move data ''
        ''''''''''''''''''''''''''''''''''''''''
        wb.SlicerCaches(slicer_name).ClearManualFilter
        ' Bottom is for totals, if you need them.
        ws.Cells(row_index, col_index).Value = ws.Range(copy_range).Value
    End With
    
End Sub

Sub Call_Slicer()
    
    ' Procedure that calls the above functions
    
    Call initialize_global_variables ' Initialize Arrays
    
    Dim wb As Workbook:  Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets(SHEET_NAME) ' Should be the sheet you want your data on
    Dim slicerItem As Variant
    Dim row_index As Integer
    
    row_index = ROW
    For Each slicerItem In SLICER_ITEM_1
        Call first_slicer_loop(wb, ws, SLICER_1, CStr(slicerItem), SLICER_ITEM_1)
        Call second_slicer_loop(wb, ws, SLICER_2, COPY_CELL, row_index, COL)
        row_index = row_index + 1
    Next slicerItem
    
    ' Bottom totals
    wb.SlicerCaches(SLICER_1).ClearManualFilter
    Call second_slicer_loop(wb, ws, SLICER_2, COPY_CELL, row_index, COL)
    
End Sub


