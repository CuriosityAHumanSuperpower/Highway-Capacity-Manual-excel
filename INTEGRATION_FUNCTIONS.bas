Attribute VB_Name = "INTEGRATION_FUNCTIONS"
Sub a_b_INTEGRATION(varX As Range, table As String, output_a As Range, output_b As Range) '2D INTERPOLATION
    'VARS
    Dim N As Integer
    N = varX.Rows.Count
    Dim a As Variant
    ReDim a(1 To N)
    Dim b As Variant
    ReDim b(1 To N)
    Dim x As Double
    'LOOP
    For i = 1 To N
        'GET VALUE
        x = varX.Cells(RowIndex:=i, ColumnIndex:=1).value
        'SET VALUES
        Worksheets(table).Range("C4").value = x
        'WAIT
        Call WaitUntilFinishedLoop
        'EVALUATE
        a(i) = Worksheets(table).Range("C7").value
        b(i) = Worksheets(table).Range("C8").value
'       MsgBox Results(i)
    Next
    output_a.value = Application.Transpose(a)
    output_b.value = Application.Transpose(b)
End Sub

Sub THREED_INTERPOLATION(varX As Range, varY As Range, table As String, output As Range) '3D INTERPOLATION
    'VARS
    Dim N As Integer
    N = varX.Rows.Count
    Dim Results As Variant
    ReDim Results(1 To N)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    'LOOP
    For i = 1 To N
        'GET VALUE
        x = varX.Cells(RowIndex:=i, ColumnIndex:=1).value
        y = varY.Cells(RowIndex:=i, ColumnIndex:=1).value
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

Sub FOURD_INTERPOLATION(varX As Range, varY As Range, varZ As Range, table As String, output As Range) '4D INTERPOLATION
    'VARS
    Dim N As Integer
    N = varX.Rows.Count
    Dim Results As Variant
    ReDim Results(1 To N)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    'LOOP
    For i = 1 To N
        'GET VALUE
        x = varX.Cells(RowIndex:=i, ColumnIndex:=1).value
        y = varY.Cells(RowIndex:=i, ColumnIndex:=1).value
        z = varZ.Cells(RowIndex:=i, ColumnIndex:=1).value
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



