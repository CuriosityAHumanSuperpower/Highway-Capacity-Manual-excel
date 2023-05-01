Attribute VB_Name = "TWO_LANE_HIGHWAY_SPECIFIC_GRADE"
'''''''''''''''''''''''''''''''''
'TWO LANE HIGHWAY_SPECIFIC GRADE'
'''''''''''''''''''''''''''''''''

'Run all
'=======


Sub TWO_LANE_HIGHWAY_SG()
    Call TWO_LANE_HIGHWAY_SG_ATS
    Call TWO_LANE_HIGHWAY_SG_PSTF
End Sub

Sub TWO_LANE_HIGHWAY_SG_ATS()
    Call VO_ATS_ITERATION
    Call VD_ATS_ITERATION
    Call fnp_ATS_D
End Sub

Sub TWO_LANE_HIGHWAY_SG_PSTF()
    Call VO_PTSF_ITERATION
    Call VD_PTSF_ITERATION
    Call fnp_PTSF_D
    Call a_b
End Sub


'Average Travel Speed (AVT)
'==========================

'Opposite direction
'------------------

Private Sub VO_ATS_ITERATION()
    'INITIALISATION
    Dim VO As Range
    Set VO = [TWO_LANE_HIGHWAY_SG[Vo_ATS_init '[pc/h']]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 5
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max 'or ([TWO_LANE_HIGHWAY_SG[Vo_ATS_'' '[pc/h']]] - [TWO_LANE_HIGHWAY_SG[Vo_ATS '[pc/h']]] <= e)
        Call fG_ATS_O_UP(VO)
        Call fG_ATS_O_DOWN(VO)
        Call ETC_ATS_O(VO)
        Call ET_ATS_O_UP(VO)
        Call ER_ATS_O_UP(VO)
        Call ET_ATS_O_DOWN(VO)
        Call ER_ATS_O_DOWN(VO)
        i = i + 1
        [TWO_LANE_HIGHWAY_SG[Vo_ATS '[pc/h']]].Copy
        [TWO_LANE_HIGHWAY_SG[Vo_ATS_'' '[pc/h']]].PasteSpecial (xlPasteValues)
        Set VO = [TWO_LANE_HIGHWAY_SG[Vo_ATS_'' '[pc/h']]]
    Wend
End Sub

Sub fG_ATS_O_UP(VO As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VO, "20-13_fg_ATS", [TWO_LANE_HIGHWAY_SG[fG_ATS_O_UP '[20-13']]])
End Sub

Sub fG_ATS_O_DOWN(VO As Range)
    Call THREED_INTERPOLATION(VO, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-7_FG_ATS", [TWO_LANE_HIGHWAY_SG[fG_ATS_O_DOWN '[20-7']]])
End Sub

Sub ETC_ATS_O(VO As Range)
    Call THREED_INTERPOLATION([GENERAL_SG[FFS - Truck scrawl speed]], VO, "20-18_ETC_ATS", [TWO_LANE_HIGHWAY_SG[ETC_ATS_O '[20-18']]])
End Sub

Sub ET_ATS_O_UP(VO As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VO, "20-15_Et_ATS", [TWO_LANE_HIGHWAY_SG[ET_ATS_O_UP '[20-15']]])
End Sub

Sub ER_ATS_O_UP(VO As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VO, "20-17_ER_ATS", [TWO_LANE_HIGHWAY_SG[ER_ATS_O_UP '[20-17']]])
End Sub

Sub ET_ATS_O_DOWN(VO As Range)
    Call THREED_INTERPOLATION(VO, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-9_ET_ATS", [TWO_LANE_HIGHWAY_SG[ET_ATS_O_DOWN '[20-9']]])
End Sub

Sub ER_ATS_O_DOWN(VO As Range)
    Call THREED_INTERPOLATION(VO, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-9_ER_ATS", [TWO_LANE_HIGHWAY_SG[ER_ATS_O_DOWN '[20-9']]])
End Sub


'Studied Direction
'------------------

Private Sub VD_ATS_ITERATION()
    'INITIALISATION
    Dim VD As Range
    Set VD = [TWO_LANE_HIGHWAY_SG[Vd_ATS_init '[pc/h']]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 5
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max 'or ([TWO_LANE_HIGHWAY_SG[Vo_ATS_'' '[pc/h']]] - [TWO_LANE_HIGHWAY_SG[Vo_ATS '[pc/h']]] <= e)
        Call fG_ATS_D_UP(VD)
        Call fG_ATS_D_DOWN(VD)
        Call ETC_ATS_D(VD)
        Call ET_ATS_D_UP(VD)
        Call ER_ATS_D_UP(VD)
        Call ET_ATS_D_DOWN(VD)
        Call ER_ATS_D_DOWN(VD)
        i = i + 1
        [TWO_LANE_HIGHWAY_SG[Vd_ATS '[pc/h']]].Copy
        [TWO_LANE_HIGHWAY_SG[Vd_ATS_'' '[pc/h']]].PasteSpecial (xlPasteValues)
        Set VD = [TWO_LANE_HIGHWAY_SG[Vd_ATS_'' '[pc/h']]]
    Wend
End Sub

Sub fG_ATS_D_UP(VD As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VD, "20-13_fg_ATS", [TWO_LANE_HIGHWAY_SG[fG_ATS_D_UP '[20-13']]])
End Sub

Sub fG_ATS_D_DOWN(VD As Range)
    Call THREED_INTERPOLATION(VD, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-7_FG_ATS", [TWO_LANE_HIGHWAY_SG[fG_ATS_D_DOWN '[20-7']]])
End Sub

Sub ETC_ATS_D(VD As Range)
    Call THREED_INTERPOLATION([GENERAL_SG[FFS - Truck scrawl speed]], VD, "20-18_ETC_ATS", [TWO_LANE_HIGHWAY_SG[ETC_ATS_D '[20-18']]])
End Sub

Sub ET_ATS_D_UP(VD As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VD, "20-15_Et_ATS", [TWO_LANE_HIGHWAY_SG[ET_ATS_D_UP '[20-15']]])
End Sub

Sub ER_ATS_D_UP(VD As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VD, "20-17_ER_ATS", [TWO_LANE_HIGHWAY_SG[ER_ATS_D_UP '[20-17']]])
End Sub

Sub ET_ATS_D_DOWN(VD As Range)
    Call THREED_INTERPOLATION(VD, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-9_ET_ATS", [TWO_LANE_HIGHWAY_SG[ET_ATS_D_DOWN '[20-9']]])
End Sub

Sub ER_ATS_D_DOWN(VD As Range)
    Call THREED_INTERPOLATION(VD, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-9_ER_ATS", [TWO_LANE_HIGHWAY_SG[ER_ATS_D_DOWN '[20-9']]])
End Sub


'fnp ATS
'-------

Private Sub fnp_ATS_D()
    Call FOURD_INTERPOLATION([TWO_LANE_HIGHWAY_SG[FFSd '[km/h']]], [TWO_LANE_HIGHWAY_SG[Vo_ATS '[pc/h']]], [INPUTS_SG[No passing zone '[%']]], "20-19_fnp_ATS", [TWO_LANE_HIGHWAY_SG[fnp_ATS '[20-19']]])
End Sub
    


'(PSTF)
'======


'Opposite direction
'------------------

Private Sub VO_PTSF_ITERATION()
    'INITIALISATION
    Dim VO As Range
    Set VO = [TWO_LANE_HIGHWAY_SG[Vo_PTSF_init '[pc/h']]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 5
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max 'or ([TWO_LANE_HIGHWAY_SG[Vo_PTSF_'' '[pc/h']]] - [TWO_LANE_HIGHWAY_SG[Vo_PTSF '[pc/h']]] <= e)
        Call fG_PTSF_O_UP(VO)
        Call fG_PTSF_O_DOWN(VO)
        Call ET_PTSF_O_UP(VO)
        Call ET_PTSF_O_DOWN(VO)
        i = i + 1
        [TWO_LANE_HIGHWAY_SG[Vo_PTSF '[pc/h']]].Copy
        [TWO_LANE_HIGHWAY_SG[Vo_PTSF_'' '[pc/h']]].PasteSpecial (xlPasteValues)
        Set VO = [TWO_LANE_HIGHWAY_SG[Vo_PTSF_'' '[pc/h']]]
    Wend
End Sub

Sub fG_PTSF_O_UP(VO As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VO, "20-14_fg_PTSF", [TWO_LANE_HIGHWAY_SG[fG_PTSF_O_UP '[20-14']]])
End Sub

Sub fG_PTSF_O_DOWN(VO As Range)
    Call THREED_INTERPOLATION(VO, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-8_FG_PTSF", [TWO_LANE_HIGHWAY_SG[fG_PTSF_O_DOWN '[20-8']]])
End Sub

Sub ET_PTSF_O_UP(VO As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VO, "20-16_ET&RV_PTSF", [TWO_LANE_HIGHWAY_SG[ET_PTSF_O_UP '[20-16']]])
End Sub

Sub ET_PTSF_O_DOWN(VO As Range)
    Call THREED_INTERPOLATION(VO, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-10_ET_PTSF", [TWO_LANE_HIGHWAY_SG[ET_PTSF_O_DOWN '[20-10']]])
End Sub


'Studied Direction
'------------------

Private Sub VD_PTSF_ITERATION()
    'INITIALISATION
    Dim VD As Range
    Set VD = [TWO_LANE_HIGHWAY_SG[Vd_PTSF_init '[pc/h']]]
    Dim i As Integer
    i = 0
    Dim i_max As Integer
    i_max = 5
    Dim e As Integer
    e = 10
    'ITERATION
    While i < i_max 'or ([TWO_LANE_HIGHWAY_SG[Vd_PTSF_'' '[pc/h']]] - [TWO_LANE_HIGHWAY_SG[Vd_PTSF '[pc/h']]] <= e)
        Call fG_PTSF_D_UP(VD)
        Call fG_PTSF_D_DOWN(VD)
        Call ET_PTSF_D_UP(VD)
        Call ET_PTSF_D_DOWN(VD)
        i = i + 1
        [TWO_LANE_HIGHWAY_SG[Vd_PTSF '[pc/h']]].Copy
        [TWO_LANE_HIGHWAY_SG[Vd_PTSF_'' '[pc/h']]].PasteSpecial (xlPasteValues)
        Set VD = [TWO_LANE_HIGHWAY_SG[Vd_PTSF_'' '[pc/h']]]
    Wend
End Sub

Sub fG_PTSF_D_UP(VD As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VD, "20-14_fg_PTSF", [TWO_LANE_HIGHWAY_SG[fG_PTSF_D_UP '[20-14']]])
End Sub

Sub fG_PTSF_D_DOWN(VD As Range)
    Call THREED_INTERPOLATION(VD, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-8_FG_PTSF", [TWO_LANE_HIGHWAY_SG[fG_PTSF_D_DOWN '[20-8']]])
End Sub


Sub ET_PTSF_D_UP(VD As Range)
    Call FOURD_INTERPOLATION([INPUTS_SG[Grade '[%']]], [INPUTS_SG[Length '[km']]], VD, "20-16_ET&RV_PTSF", [TWO_LANE_HIGHWAY_SG[ET_PTSF_D_UP '[20-16']]])
End Sub

Sub ET_PTSF_D_DOWN(VD As Range)
    Call THREED_INTERPOLATION(VD, [TWO_LANE_HIGHWAY_SG[Type Terrain]], "20-10_ET_PTSF", [TWO_LANE_HIGHWAY_SG[ET_PTSF_D_DOWN '[20-10']]])
End Sub


'fnp PTSF
'--------

Private Sub fnp_PTSF_D()
    Call FOURD_INTERPOLATION([TWO_LANE_HIGHWAY_SG[FFSd '[km/h']]], [TWO_LANE_HIGHWAY_SG[Vd_PTSF '[pc/h']]], [INPUTS_SG[No passing zone '[%']]], "20-20_fnp_PTSF", [TWO_LANE_HIGHWAY_SG[fnp_PTSF '[20-20']]])
End Sub
 
'a & b
'-----

Private Sub a_b()
    Call a_b_INTEGRATION([TWO_LANE_HIGHWAY_SG[Vo_PTSF '[pc/h']]], "20-21_a&b", [TWO_LANE_HIGHWAY_SG[a '[20-21']]], [TWO_LANE_HIGHWAY_SG[b '[20-21']]])
End Sub


