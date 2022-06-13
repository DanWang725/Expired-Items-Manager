Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="22,23,24,25"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveCell.SpecialCells(xlCellTypeSameValidation).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("B2:G17").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
End Sub

Sub YearColour()
Attribute YearColour.VB_ProcData.VB_Invoke_Func = " \n14"
'
' YearColour Macro
'

'
    Columns("I:I").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub
Sub YearColourRemove()
Attribute YearColourRemove.VB_ProcData.VB_Invoke_Func = " \n14"
'
' YearColourRemove Macro
'

'
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("I:I").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    Selection.ClearContents
End Sub
Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = " \n14"
'
' test Macro
'

'
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=SMALL(RC[1]:RC[6],1)"
    Range("E7").Select
End Sub
