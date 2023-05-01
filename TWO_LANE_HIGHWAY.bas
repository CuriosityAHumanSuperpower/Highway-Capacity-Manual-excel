Attribute VB_Name = "TWO_LANE_HIGHWAY"
''''''''''''''''''
'TWO LANE HIGHWAY'
''''''''''''''''''


'Run all
'=======
 
Sub RUN_PISTA_SIMPLES()
    Call VP_ATS
    Call FNP_INTERPOLATION
    Call VP_PTSF
    Call FDNP_INTERPOLATION
End Sub



'Average Travel Speed (AVT)
'==========================

Sub VP_ATS()
    'INITIALISATION
    Dim VP As Range
    Set VP = [TWO_LANE_HIGHWAY_G[vp1_init]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 9
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max
        i = i + 1
        [TWO_LANE_HIGHWAY_G[vp1]].Copy
        [TWO_LANE_HIGHWAY_G[vp1'']].PasteSpecial (xlPasteValues)
        Set VO = [TWO_LANE_HIGHWAY_G[vp1'']]
    Wend
End Sub

'(PSTF)
'======

Sub VP_PTSF()
    'INITIALISATION
    Dim VP As Range
    Set VP = [TWO_LANE_HIGHWAY_G[vp2_init]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 9
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max
        i = i + 1
        [TWO_LANE_HIGHWAY_G[vp2]].Copy
        [TWO_LANE_HIGHWAY_G[vp2'']].PasteSpecial (xlPasteValues)
        Set VO = [TWO_LANE_HIGHWAY_G[vp2'']]
    Wend
End Sub


'Other
'=====

Private Function MAX_MAX(rangeA As Range, rangeB As Range) As Range
    Dim N As Integer
    Dim M As Integer
    N = rangeA.Rows.Count
    M = rangeB.Rows.Count
    If N = M Then
        Dim Results As Variant
        ReDim Results(1 To N)
        For i = 1 To N
            Results(i) = Application.Max(rangeA(i), rangeB(i))
        Next
    End If
    MAX_MAX = Application.Transpose(Results)
End Function

Sub FDNP_INTERPOLATION() '4D INTERPOLATION
    'INPUT
    Dim matrix As Range
'    Set matrix = Range([INTERP_INPUTS_FDNP].value)
    Set matrix = [TWO_LANE_HIGHWAY_G[[X'']:[Table'']]]
    'OUTPUT
    Dim output As Range
'    Set output = Range([INTERP_OUTPUTS_FDNP].value)
    Set output = [TWO_LANE_HIGHWAY_G[fd/np'']]
    'VARS
    Dim N As Integer
    Dim M As Integer
    N = matrix.Rows.Count
    M = matrix.Columns.Count
    Dim Results As Variant
    ReDim Results(1 To N)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    'LOOP
    For i = 1 To N
        'GET VALUE
        x = matrix.Cells(RowIndex:=i, ColumnIndex:=1).value
        y = matrix.Cells(RowIndex:=i, ColumnIndex:=2).value
        z = matrix.Cells(RowIndex:=i, ColumnIndex:=3).value
        table = matrix.Cells(RowIndex:=i, ColumnIndex:=4).value
        'SET VALUES
        Worksheets(table).Range("C3").value = x
        Worksheets(table).Range("C4").value = y
        Worksheets(table).Range("C5").value = z
        'WAIT
        Call WaitUntilFinishedLoop
        'EVALUATE
        Results(i) = Worksheets(table).Range("C8").value
'        MsgBox Results(i)
    Next
    output.value = Application.Transpose(Results)
End Sub

Sub FNP_INTERPOLATION() '3D INTERPOLATION
    'INPUT
    Dim matrix As Range
'    Set matrix = Range([INTERP_INPUTS_FNP].value)
    Set matrix = [TWO_LANE_HIGHWAY_G[[X]:[Table]]]
    'OUTPUT
    Dim output As Range
'    Set output = Range([INTERP_OUTPUTS_FNP].value)
    Set output = [TWO_LANE_HIGHWAY_G[fnp'']]
    'VARS
    Dim N As Integer
    Dim M As Integer
    N = matrix.Rows.Count
    M = matrix.Columns.Count
    Dim Results As Variant
    ReDim Results(1 To N)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    'LOOP
    For i = 1 To N
        'GET VALUE
        x = matrix.Cells(RowIndex:=i, ColumnIndex:=1).value
        y = matrix.Cells(RowIndex:=i, ColumnIndex:=2).value
        table = matrix.Cells(RowIndex:=i, ColumnIndex:=3).value
        'WAIT
        Call WaitUntilFinishedLoop
        'SET VALUES
        Worksheets(table).Range("C3").value = x
        Worksheets(table).Range("C4").value = y
        'WAIT
        Call WaitUntilFinishedLoop
        'EVALUATE
        Results(i) = Worksheets(table).Range("H3").value
'       MsgBox Results(i)
    Next
    output.value = Application.Transpose(Results)
End Sub

