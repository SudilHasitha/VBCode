Sub DailyRentalAdapter14012020v5()
'
' DailyRentalAdapter14012020v5 Macro

' Check Criteria Array and add values if needed
' Change the name of worksheet

    Dim Worksheet As String
    Worksheet = InputBox("Current Workbook ")


'
    Selection.AutoFilter
    '"$A$1:$K$477" should replace with all ctrl+A
    '"ib_null_null", "Ib_ppp_ttt", "ib_www_qqq", "ib_xxx_zzz", "TestApp1", "TestRental", "ib_xxx_www", "ib_rnt_wrnt", "ib_www_rrr", "ib_www_eee",_
    '"ib_alt_day"
    ActiveSheet.Range("A:K").AutoFilter Field:=9, Criteria1:=Array( _
        "ib_null_null", "Ib_ppp_ttt", "ib_www_qqq", "ib_xxx_zzz", "TestApp1", "TestRental", "ib_xxx_www", "ib_rnt_wrnt", "ib_www_rrr", "ib_www_eee") _
        , Operator:=xlFilterValues
        
    '"A46" should replace with cell below the A1
    'first select A1 as active cell
    Range("A1").Select
    'Range("A46").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    
    Selection.AutoFilter
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$L$1000").AutoFilter Field:=10, Criteria1:="0"
    
    ' A3 - should replace with cell immediate below to A1
    'Range("A3").Select
    Range("A1").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    
    Selection.AutoFilter
    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("L2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(8, 9)), TrailingMinusNumbers:=True
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("A1").Select
    ActiveWorkbook.Worksheets(Worksheet).Sort.SortFields.Clear
    
    '"B2:B229" B229 should replace with last value in Column B
    Range("B2").Select
    Dim lastRow As Integer
    lastRow = Cells.Find(What:="*", _
                    After:=Range("B2"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    'end
    ActiveWorkbook.Worksheets(Worksheet).Sort.SortFields.Add Key _
        :=Range("B2:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    
    '"A1:L229" - should replace with ctrl+A
    With ActiveWorkbook.Worksheets(Worksheet).Sort
        .SetRange Range("A:L")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Subtotal GroupBy:=2, Function:=xlCount, TotalList:=Array(2), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    ActiveSheet.Outline.ShowLevels RowLevels:=2
End Sub
