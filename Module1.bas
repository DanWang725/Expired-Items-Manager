Attribute VB_Name = "Module1"
Sub MakeDates()

    Dim day As String
    Dim month As String
    Dim year As String
    Dim lCol As Integer
    Dim final As String
    Application.ScreenUpdating = False
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim cell As Range
    For Each cell In Range(Cells(2, 1), Cells(lCol, 1))
        If Not IsEmpty(cell.Offset(0, 9).Value) Then
            If Not IsEmpty(cell.Offset(0, 10).Value) Then
                day = CStr(cell.Offset(0, 10).Value)
            Else
                day = "1"
            End If
            
            
            month = cell.Offset(0, 9).Value
            
            If Not IsEmpty(cell.Offset(0, 8).Value) Then
                year = CStr(20 & cell.Offset(0, 8))
            Else
                year = CStr(2022)
            End If
            
            final = month & " " & day & ", " & year
            cell.Offset(0, 7).Value = final
        End If
            
        Range(cell.Offset(0, 8), cell.Offset(0, 10)).Select
        Selection.ClearContents
    Next cell
    
    Call RemoveFormat
    Call DeColourAll
    
    Call ShiftAll
    Call ReapplyFormat
    Call ColourAll
    Application.ScreenUpdating = True

End Sub

Sub SortDatesForItem()
'
' SortAllAddedDates Macro
'

'
    Range("C2:F2").Select
    Selection.Copy
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("M2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("M2:M5")
        .Header = xlNo
        .MatchCase = True
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
End Sub

Sub RemoveFormat()
    Dim lCol As Integer
    
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(2, 9), Cells(lCol, 10)).Select
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


Sub ReapplyFormat()
    Dim lCol As Integer
    
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range(Cells(2, 9), Cells(lCol, 9)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="22,23,24,25,26,27,28,29,30"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    Range(Cells(2, 10), Cells(lCol, 10)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="May,June,July,August,September,October,November,December,January,February,March,April"
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

Sub ShiftAll()
'
' Macro3 Macro
'

'
    Dim lCol As Integer
    
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range(Cells(2, 1), Cells(lCol, 8)).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    
    Dim cell As Range
    For Each cell In Range(Cells(2, 2), Cells(lCol, 2))
        cell.Select
        ActiveCell.FormulaR1C1 = "=SMALL(RC[1]:RC[6],1)"
        
        ' sort all the dates
        If Not IsEmpty(cell.Offset(0, 1).Value) Then
            Range(cell.Offset(0, 1), cell.Offset(0, 6)).Select
            Selection.Copy
            Range("M2").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True
            Application.CutCopyMode = False
            ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("M2"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("Sheet1").Sort
                .SetRange Range("M2:M7")
                .Header = xlNo
                .MatchCase = True
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        
            Selection.Copy
            Range(cell.Offset(0, 1), cell.Offset(0, 6)).Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True
        End If
        
    Next cell
    Range("M2:M7").Clear
    
End Sub

Sub DebugDelete()
    Selection.Delete
End Sub

Sub DeleteSelectedExpiryDates()
    Dim sel As Range
    Dim cs As Range
    Dim valid As Boolean
    Dim selCoords As Integer
    
    valid = True
    Set sel = Selection
    
    For Each cs In sel.Cells
        If cs.Row = 1 Or cs.Column <= 2 Or cs.Column >= 8 Then
            valid = False
        End If
        selCoords = cs.Row
    Next
    
    If valid Then
        Application.ScreenUpdating = False
        Call RemoveFormat
        Call DeColourAll
    
        sel.Delete Shift:=xlToLeft
        
        Call ReapplyFormat
        Call ColourAll
        
        Cells(selCoords, 3).Select
        Application.ScreenUpdating = True
    Else
        MsgBox "INVALID SELECTION: The selection contains cells that are not expiry dates"
    End If
End Sub

Sub RemoveShiftSelection() 'depracated function
    Dim selRange As Range
    Set selRange = Selection
    
    If Selection.Cells.Count = 1 Then
        If selRange.Row > 1 And selRange.Column > 2 And selRange.Column < 8 Then
    
            Call RemoveFormat
            Call DeColourAll
    
            selRange.Select
            Selection.Delete Shift:=xlToLeft
    
            Call ReapplyFormat
            Call ColourAll
        Else
            MsgBox "Please select a date to delete"
        End If
    End If
End Sub

Sub ColourAll()
'
' ColourAll Macro
'

'
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(2, 9), Cells(lCol, 9)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    Range(Cells(2, 10), Cells(lCol, 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range(Cells(2, 11), Cells(lCol, 11)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range(Cells(2, 9), Cells(lCol, 11)).Select
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
End Sub
Sub DeColourAll()
'
' DeColourAll Macro
'

'
    lCol = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(2, 9), Cells(lCol, 11)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
