Attribute VB_Name = "Module3"
Sub SortExpNoNum()
Attribute SortExpNoNum.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortExpNoNum Macro
'

'
    ActiveSheet.Range("$B$1:$B$137").AutoFilter Field:=1, Operator:= _
        xlFilterValues, Criteria2:=Array(0, "10/12/2026", 0, "1/31/2022")
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B137"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B137"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SortToEarlist()
Attribute SortToEarlist.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortToEarlist Macro
'

'
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B137"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$K$137").AutoFilter Field:=2, Operator:= _
        xlFilterValues, Criteria2:=Array(0, "4/25/2024", 0, "7/31/2023", 0, "11/6/2022")
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B137"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
End Sub
Sub ShowEarliest()
Attribute ShowEarliest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ShowEarliest Macro
'

'
    ActiveSheet.Range("$A$1:$K$137").AutoFilter Field:=2, Operator:= _
        xlFilterValues, Criteria2:=Array(0, "4/25/2024", 0, "7/31/2023", 0, "11/6/2022")
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B137"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ShowAll()
Attribute ShowAll.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ShowAll Macro
'

'
    ActiveSheet.Range("$A$1:$K$137").AutoFilter Field:=2
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("A1:A137"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
