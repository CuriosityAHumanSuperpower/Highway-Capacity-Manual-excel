Attribute VB_Name = "YEAR_BY_YEAR"
''''''''
'INPUTS'
''''''''

Sub INITIALIZATION()
    'ERASE DATA
    Call DEL_TABLE_ROWS("YEAR_BY_YEAR", "YEAR_BY_YEAR")
    'ERASE OUTPUT COLUMNS
    Call ERASE_OUTPUT
    'SET THE IDs
    [INPUT[Id]].Copy
    [YEAR_BY_YEAR[ID]].PasteSpecial (xlPasteValues)
End Sub

Sub RUN_YEAR_BY_YEAR()
    'VAR
    Set yearTable = ActiveSheet.ListObjects("YEAR_BY_YEAR")
    Set inputTable = Worksheets("INPUTS").ListObjects("INPUT")
    Dim N As Integer
    N = yearTable.ListColumns.Count
    Dim year As String
    Dim k As Integer
    Dim output As Variant
    output = Array("LOS_", "ATS_", "PTSF_", "VP_", "D_", "S_")
    Dim M As Integer
    M = UBound(output, 1)
    Dim i As Integer
    i = N
    'ERASE OUTPUT COLUMNS
    Call ERASE_OUTPUT
    'ITERATION OVER THE YEARS
    k = 2
    N = yearTable.ListColumns.Count
    While k < 1 + N * M
        year = yearTable.ListColumns(k).Name
        'TRANSFERT DATA TO "INPUTS" SHEET
        yearTable.ListColumns(k).DataBodyRange.Copy
        [INPUT[Total '[VDMA']]].PasteSpecial (xlPasteValues)
        'RUN CODE
        Call RUN_HCM2000
        'ADD OUTPUT COLUMNS
        For Each colName In output
            k = k + 1
            Set newCol = yearTable.ListColumns.Add(Position:=k)
            newCol.Name = colName & year
            yearTable.HeaderRowRange(k).Interior.ColorIndex = 1
        Next
        k = k + 1
        'GET THE RESULTS
        For Each colName In output
            inputTable.ListColumns(colName).DataBodyRange.Copy
            yearTable.ListColumns(colName & year).DataBodyRange.PasteSpecial (xlPasteValues)
        Next
   Wend
End Sub

Sub ERASE_OUTPUT()
    'ERASE OUTPUT COLUMNS
    Set yearTable = ActiveSheet.ListObjects("YEAR_BY_YEAR")
    Dim i As Integer
    Dim N As Integer
    N = yearTable.ListColumns.Count
    i = N
    While i > 1
        colName = yearTable.ListColumns(i).Name
        If InStr(colName, "_") > 0 Then
            yearTable.ListColumns(i).Delete
        End If
        i = i - 1
    Wend
End Sub
