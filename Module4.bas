Attribute VB_Name = "Module4"
Sub SortByDrinks()
Attribute SortByDrinks.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortByDrinks Macro
'

'
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add(Range( _
        "A1:A497"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(189, 215, 238)
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SortAllAddedDates()
Attribute SortAllAddedDates.VB_ProcData.VB_Invoke_Func = " \n14"
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
