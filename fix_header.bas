Sub Fix_header()
    Dim main_file As Workbook
    Dim main_sheet As Worksheet
    Set main_file = ActiveWorkbook
    Set main_sheet = main_file.ActiveSheet
    
    lastCol_1 = main_sheet.Cells(1, Columns.Count).End(xlToLeft).Column
    lastCol_2 = main_sheet.Cells(2, Columns.Count).End(xlToLeft).Column
    If lastCol_1 > lastCol_2 Then
        lastCol = lastCol_1
    Else
        lastCol = lastCol_2
    End If
    main_sheet.Rows(3).EntireRow.Insert
    main_sheet.Rows(4).EntireRow.Insert
    main_sheet.Cells(3, 1) = main_sheet.Cells(1, 1)
    For c = 2 To lastCol
        If main_sheet.Cells(1, c) = "" Then
            main_sheet.Cells(3, c) = main_sheet.Cells(3, c - 1)
        Else
            main_sheet.Cells(3, c) = main_sheet.Cells(1, c)
        End If
    Next
    For c = 1 To lastCol
        If main_sheet.Cells(2, c) = "" Then
            main_sheet.Cells(4, c) = main_sheet.Cells(3, c)
        Else
            main_sheet.Cells(4, c) = main_sheet.Cells(3, c) & " | " & main_sheet.Cells(2, c)
        End If
    Next
    main_sheet.Range("1:3").EntireRow.Delete
    main_sheet.Range("1:1").WrapText = False
    
    main_sheet.Rows("2:" & main_sheet.Rows.Count).Hidden = True
    For Each col In main_sheet.UsedRange.Columns
        col.AutoFit
        col.ColumnWidth = col.ColumnWidth
    Next col
    main_sheet.Rows("2:" & main_sheet.Rows.Count).Hidden = False
    main_sheet.Rows(1).AutoFit
End Sub
