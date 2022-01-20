Attribute VB_Name = "MiseEnForme"
Function GridLineColor(value As Integer)
Attribute GridLineColor.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWindow.GridlineColorIndex = value
End Function
Function FontColor(rng, PatternColorIndex, ThemeColor) 'xlAutomatic,xlThemeColorAccent2 ; change la couleur de font
Attribute FontColor.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.CutCopyMode = False
    With rng.Interior
        .PatternColorIndex = PatternColorIndex
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Function
Function GridColor(rng, LineStyle, ThemeColor, Weight) 'xlContinuous,6,xlThin ; change la couleur des borudre
Attribute GridColor.VB_ProcData.VB_Invoke_Func = " \n14"

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = LineStyle
        .ThemeColor = ThemeColor
        .TintAndShade = 0
        .Weight = Weight
    End With
End Function

Function BackColorNone(rng) ' met la couleur de fond en blanc
Attribute BackColorNone.VB_ProcData.VB_Invoke_Func = " \n14"
     With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Function
