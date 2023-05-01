Attribute VB_Name = "INPUTS"
''''''''
'INPUTS'
''''''''

'STEPS:
'1. GET ALL THE INFORMATIONS FROM THE TABLE "INPUT"
'2. GIVEN THE METHODE (TWO LANE, MULTILANE WITH OR WITHOUT SPECIFIC GRADE),
'   THE INPUTS WILL BE COPIED AND PASTE TO THE APPROPRIATE TABLE
'3. ID ARE UNIQUE
'4. FOR EACH METHODE, THE eLOS VALUES ARE ESTIMATED
'5. THE VALUES ARE PUSHED BACK IN THE "INPUT" TABLE USING IDs TO JOIN THE RESULTS TO APPROPRIATE INPUT

Sub RUN_HCM2000()
    'DESABLE SCREEN
    Call TurnEverythingOff
    'VAR
    Dim INDICATOR_TWO_LANE As Integer
    INDICATOR_TWO_LANE = 0
    Dim INDICATOR_TWO_LANE_SG As Integer
    INDICATOR_TWO_LANE_SG = 0
    Dim INDICATOR_MULTI_LANE As Integer
    INDICATOR_MULTI_LANE = 0
    Dim INDICATOR_MULTI_LANE_SG As Integer
    INDICATOR_MULTI_LANE_SG = 0
    Set wsINPUT = Worksheets("INPUTS")
    Set tableINPUT = wsINPUT.ListObjects("INPUT")
    Dim SHEETS As Variant
    SHEETS = Array("TWO LANE HIGHWAY", "TWO LANE HIGHWAY_SPECIAL GRADE", "MULTILANE HIGHWAY", "MULTILANE HIGHWAY_SPECIAL GRADE")
    Dim TABLES As Variant
    TABLES = Array("TWO_LANE_HIGHWAY_G", "INPUTS_SG", "MULTILANE_HIGHWAY", "MULTILANE_HIGHWAY_SPECIAL_GRADE")
    Dim M As Integer
    Dim ws As Worksheet
    Dim table As ListObject
    Dim newrow As ListRow
    Dim IDs As Range
    Set IDs = [INPUT[Id]]
    Dim MODELs As Range
    Set MODELs = [INPUT[Modelo]]
    Dim N As Integer
    N = IDs.Rows.Count
    Set Results = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    'DEL PREVIOUS VALUES
    Call DEL_TABLE_ROWS("TWO LANE HIGHWAY_SPECIAL GRADE", "TWO_LANE_HIGHWAY_SG")
    Call DEL_TABLE_ROWS("TWO LANE HIGHWAY_SPECIAL GRADE", "GENERAL_SG")
    Call DEL_TABLE_ROWS("TWO LANE HIGHWAY_SPECIAL GRADE", "INPUTS_SG")
    Call DEL_TABLE_ROWS("TWO LANE HIGHWAY", "TWO_LANE_HIGHWAY_G")
    Call DEL_TABLE_ROWS("MULTILANE HIGHWAY", "MULTILANE_HIGHWAY")
    Call DEL_TABLE_ROWS("MULTILANE HIGHWAY_SPECIAL GRADE", "MULTILANE_HIGHWAY_SPECIAL_GRADE")
    'LOOP: PUSH TO APPROPRIATE TABLE
    For i = 1 To N
        ID = IDs.Cells(RowIndex:=i, ColumnIndex:=1).value
        model = MODELs.Cells(RowIndex:=i, ColumnIndex:=1).value
        If model = "TWO LANE HIGHWAY" Then
            INDICATOR_TWO_LANE = 1
            Set ws = Worksheets("TWO LANE HIGHWAY")
            Set table = ws.ListObjects("TWO_LANE_HIGHWAY_G")
            'ADD ROW
            Set newrow = table.ListRows.Add
            'ADD VALUES
            With newrow
                .Range(table.ListColumns("ID").Index) = ID
                .Range(table.ListColumns("Extensão").Index) = [INPUT[Extensão (km)]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Número de Faixas").Index) = [INPUT[Número de Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura das Faixas").Index) = [INPUT[Largura das Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura Acost.").Index) = [INPUT[Largura Acost.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("passeio").Index) = [INPUT[Cars '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("pesado").Index) = [INPUT[Trucks '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("rural/urbano").Index) = [INPUT[Rural/ Urban]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("plano/ondulado").Index) = [INPUT[Type of terrain]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PRv").Index) = [INPUT[PR '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("acessos/km").Index) = [INPUT[Access points /km '[1']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("BFFS").Index) = [INPUT[BFFS  '[km/h']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("% zonas s/ultra").Index) = [INPUT[No passing zone '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PHF").Index) = [INPUT[PHF]].Cells(RowIndex:=i, ColumnIndex:=1).value
            End With
        End If
        If model = "TWO LANE HIGHWAY_SPECIAL GRADE" Then
            INDICATOR_TWO_LANE_SG = 1
            Set ws = Worksheets("TWO LANE HIGHWAY_SPECIAL GRADE")
            Set table = ws.ListObjects("INPUTS_SG")
            Set table2 = ws.ListObjects("GENERAL_SG")
            Set table3 = ws.ListObjects("TWO_LANE_HIGHWAY_SG")
            'ADD ROW
            Set newrow = table.ListRows.Add
            Set newrow2 = table2.ListRows.Add
            Set newrow3 = table3.ListRows.Add
            'ADD VALUES
            With newrow
                .Range(table.ListColumns("ID").Index) = ID
                .Range(table.ListColumns("Grade [%]").Index) = [INPUT[Declividade]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Length [km]").Index) = [INPUT[Extensão (km)]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Direction analysed ?").Index) = [INPUT[Direction analysed ?]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Type of terrain_'").Index) = [INPUT[Type of terrain]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Lane width [m]").Index) = [INPUT[Largura das Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Shoulder width [m]").Index) = [INPUT[Largura Acost.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Access points /km [1]").Index) = [INPUT[Access points /km '[1']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("No passing zone [%]").Index) = [INPUT[No passing zone '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Cars [VDMA]").Index) = [INPUT[Cars '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Trucks [VDMA]").Index) = [INPUT[Trucks '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Direction Split [%]").Index) = [INPUT[Direction Split '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Truck crawl speed [km/h]").Index) = [INPUT[Truck crawl speed '[km/h']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PTC [%]").Index) = [INPUT[PTC '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PR [%]").Index) = [INPUT[PR '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PHF").Index) = [INPUT[PHF]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("BFFS  [km/h]").Index) = [INPUT[BFFS  '[km/h']]].Cells(RowIndex:=i, ColumnIndex:=1).value
            End With
        End If
        If model = "MULTILANE HIGHWAY" Then
            INDICATOR_MULTI_LANE = 1
            Set ws = Worksheets("MULTILANE HIGHWAY")
            Set table = ws.ListObjects("MULTILANE_HIGHWAY")
            'ADD ROW
            Set newrow = table.ListRows.Add
            'ADD VALUES
            With newrow
                .Range(table.ListColumns("ID").Index) = ID
                .Range(table.ListColumns("Extensão").Index) = [INPUT[Extensão (km)]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Número de Faixas").Index) = [INPUT[Número de Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura das Faixas").Index) = [INPUT[Largura das Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura Acost.").Index) = [INPUT[Largura Acost.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Separador").Index) = [INPUT[Separador]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Num. Faixas Marg.").Index) = [INPUT[Num. Faixas Marg.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("distrib. direc.").Index) = [INPUT[Direction Split '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("acessos/km").Index) = [INPUT[Access points /km '[1']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("BFFS").Index) = [INPUT[BFFS  '[km/h']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("% zonas s/ultra").Index) = [INPUT[No passing zone '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("passeio").Index) = [INPUT[Cars '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("pesado").Index) = [INPUT[Trucks '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("rural/urbano").Index) = [INPUT[Rural/ Urban]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("plano/ondulado").Index) = [INPUT[Type of terrain]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PRv").Index) = [INPUT[PR '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PHF").Index) = [INPUT[PHF]].Cells(RowIndex:=i, ColumnIndex:=1).value
            End With
        End If
        If model = "MULTILANE HIGHWAY_SPECIAL GRADE" Then
            INDICATOR_MULTI_LANE_SG = 1
            Set ws = Worksheets("MULTILANE HIGHWAY_SPECIAL GRADE")
            Set table = ws.ListObjects("MULTILANE_HIGHWAY_SPECIAL_GRADE")
            'ADD ROW
            Set newrow = table.ListRows.Add
            'ADD VALUES
            With newrow
                .Range(table.ListColumns("ID").Index) = ID
                .Range(table.ListColumns("Extensão").Index) = [INPUT[Extensão (km)]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Número de Faixas").Index) = [INPUT[Número de Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Direction analysed ?").Index) = [INPUT[Direction analysed ?]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Grade [%]").Index) = [INPUT[Declividade]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura das Faixas").Index) = [INPUT[Largura das Faixas]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Largura Acost.").Index) = [INPUT[Largura Acost.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Separador").Index) = [INPUT[Separador]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("Num. Faixas Marg.").Index) = [INPUT[Num. Faixas Marg.]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("distrib. direc.").Index) = [INPUT[Direction Split '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("acessos/km").Index) = [INPUT[Access points /km '[1']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("BFFS").Index) = [INPUT[BFFS  '[km/h']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("% zonas s/ultra").Index) = [INPUT[No passing zone '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("passeio").Index) = [INPUT[Cars '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("pesado").Index) = [INPUT[Trucks '[VDMA']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("rural/urbano").Index) = [INPUT[Rural/ Urban]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("plano/ondulado").Index) = [INPUT[Type of terrain]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PRv").Index) = [INPUT[PR '[%']]].Cells(RowIndex:=i, ColumnIndex:=1).value
                .Range(table.ListColumns("PHF").Index) = [INPUT[PHF]].Cells(RowIndex:=i, ColumnIndex:=1).value
            End With
        End If
    Next
    'RUN TABLES
    If INDICATOR_TWO_LANE > 0 Then
        Call RUN_PISTA_SIMPLES
    End If
    If INDICATOR_TWO_LANE_SG > 0 Then
        Call TWO_LANE_HIGHWAY_SG
    End If
    If INDICATOR_MULTI_LANE > 0 Then
        'NO EXISTING MACRO
    End If
    If INDICATOR_MULTI_LANE_SG > 0 Then
        Call RUN_MULTILANE_SG
    End If
    'RETURN RESULTS
    For i = 0 To 3
        Set ws = Worksheets(SHEETS(i))
        Set table = ws.ListObjects(TABLES(i))
        M = table.DataBodyRange.Rows.Count
        For j = 2 To M
            idINPUT = table.DataBodyRange.Cells(j, table.ListColumns("ID").Index).value
            RowIndex = GET_ROW_INDEX("INPUTS", "INPUT", "Id", idINPUT)
            table.DataBodyRange.Cells(j, table.ListColumns("LOS_").Index).Copy
            tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("LOS_").Index).PasteSpecial (xlPasteValues)
            If i < 2 Then 'TWO LANE HIGHWAY
                table.DataBodyRange.Cells(j, table.ListColumns("ATS_").Index).Copy
                tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("ATS_").Index).PasteSpecial (xlPasteValues)
                table.DataBodyRange.Cells(j, table.ListColumns("PTSF_").Index).Copy
                tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("PTSF_").Index).PasteSpecial (xlPasteValues)
            Else 'MULTILANE HIGHWAY
                table.DataBodyRange.Cells(j, table.ListColumns("VP_").Index).Copy
                tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("VP_").Index).PasteSpecial (xlPasteValues)
                table.DataBodyRange.Cells(j, table.ListColumns("D_").Index).Copy
                tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("D_").Index).PasteSpecial (xlPasteValues)
                table.DataBodyRange.Cells(j, table.ListColumns("S_").Index).Copy
                tableINPUT.DataBodyRange.Cells(RowIndex, tableINPUT.ListColumns("S_").Index).PasteSpecial (xlPasteValues)
            End If
        Next
    Next
    'ENABLE SCREEN
    Call TurnEverythingOn
    'SAVE WORK
    ActiveWorkbook.Save
End Sub


Sub DEL_TABLE_ROWS(worksheetName As String, tableName As String)
    Dim M As Integer
    Dim ws As Worksheet
    Set ws = Worksheets(worksheetName)
    Set table = ws.ListObjects(tableName)
    M = table.DataBodyRange.Rows.Count
    For j = 2 To M '/!\ DO NOT DEL THE FIRST ROW THAT CONTAIN ALL THE FORMULATIONS
       table.ListRows(2).Delete
    Next
End Sub

Private Sub TurnEverythingOff()
    With Application
'        .Calculation = xlCalculationManual
'        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .StatusBar = "Calculations in progress"
    End With
End Sub

Private Sub TurnEverythingOn()
    With Application
'        .Calculation = xlCalculationAutomatic
'       .EnableEvents = True
        .DisplayAlerts = True
        .ScreenUpdating = True
        .StatusBar = False
    End With
End Sub

Private Function GET_ROW_INDEX(worksheetName As String, tableName As String, columnName As String, value As Variant) As Integer
    Dim M As Integer
    Dim ws As Worksheet
    Set ws = Worksheets(worksheetName)
    Set table = ws.ListObjects(tableName)
    M = table.DataBodyRange.Rows.Count
    Dim i As Integer
    i = 0
    While table.DataBodyRange.Cells(i, table.ListColumns(columnName).Index) <> value
        i = i + 1
    Wend
    GET_ROW_INDEX = i
End Function


Sub WaitUntilFinishedLoop()
'Loop until all your calculations are done
Application.Calculate 'Optional - recalculates all formulas
Do Until Application.CalculationState = xlDone
    DoEvents
Loop
'~~> Rest of your code goes here
End Sub
