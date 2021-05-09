Option Explicit
Global OverUseSMS As Long
Global Revenue As Variant


Sub VirakesiriPayment20200130PaymentV2()
'
'
'

'Set the first Payment as first active work sheet
 If ActiveSheet.Name <> "Payment" Then
    Sheets("Payment").Select
 End If


'Selection of filter
    Selection.AutoFilter
    Application.Goto Reference:="R1C25"
    ActiveSheet.Range("A:AA").AutoFilter Field:=25, Criteria1:= _
        "=FAILED", Operator:=xlOr, Criteria2:="=INSUFFICIENT BALANCE"

'first select A1 as active cell
    Range("A1").Select
    ActiveCell.Offset(1, 0).Select

'Selection of unsuccessful payment data of all payments and deletion
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    

    Selection.AutoFilter

'Remove unwanted columns
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    
 'Select F1 cell and name it as Status
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Status"
    
 'Select row 1 and change format
    Rows("1:1").Select
    Selection.Font.Bold = True
    
 'Auto fit date time column to it column heading size
    Columns("D:D").EntireColumn.AutoFit

'Select cell g1 and name it as Date
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Date"

 'select D2 and select all the records
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
 'Using text to column populate Column G with Date
    Selection.TextToColumns Destination:=Range("G2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 9), Array(3, 9)), TrailingMinusNumbers:=True
 
    Range("A1").Select
    
    
    ' Get range of data in the work sheet1
    Dim rng As Range
    Dim rngData As Range
    
    
    Set rng = ActiveSheet.Range("A1")
    Set rngData = Range(rng, rng.End(xlToRight))
    Set rngData = Range(rngData, rngData.End(xlDown))
    

    
  'Add sheet
'   Sheets.Add
    Sheets.Add(After:=ActiveSheet).Name = "Virakesari_Payment_Total"
   
   Dim PTCache As PivotCache
   Dim PT As PivotTable
   
   'Create the Cache
    Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=rngData)
    'Select the destination sheet
    Sheets("Virakesari_Payment_Total").Select
    Cells(3, 1).Select
    
    'Create the Pivot table
    Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, _
        TableDestination:=Range("A3"), TableName:="Pivot1")
    
    ActiveWorkbook.ShowPivotTableFieldList = True
    
    'Adding orderby values
    With ActiveSheet.PivotTables("Pivot1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("Pivot1").PivotFields("msisdn")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    'Add aggregation function
    ActiveSheet.PivotTables("Pivot1").AddDataField ActiveSheet.PivotTables( _
        "Pivot1").PivotFields("charged_amount"), "Sum of charged_amount", xlSum
    
    
   Dim lastRow As Integer
     lastRow = Range("A4").CurrentRegion.Rows.Count
     
    'Collapse the detailed view
     Range("A4").Select
        
    ActiveCell.End(xlDown).Select
    lastRow = lastRow - 1
    Range("A4:A" & lastRow).Select
    Selection.ShowDetail = False
    Selection.End(xlUp).Select
    
    'Get the last row after collapsing
    Range("A3").Select
     lastRow = Range("A3").CurrentRegion.Rows.Count
     
    'Get the sum of revenue and assign to global variable
        Revenue = Round(Range("B" & lastRow + 2).Value, 2)
        Debug.Print Revenue
End Sub


Sub VirakesiriSMSOutbound31012020V2()
'
' RupavahiniSMSOutbound31012020V1 Macro
'
'

    
    'Go back to SMSv2
    Worksheets("SMSv2").Activate
                    
    Dim lr As Long
    lr = Range("A1").CurrentRegion.Rows.Count
    
    'Go back to SMSv1
    Worksheets("SMSv1").Activate
    'Name N1 as client corelation coefficient
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Client_Coorelation"
    
    'Enter  index match
    '
    Range("N2").Select
    Range("N2").Formula = "=INDEX(SMSv2!$A$2:$K$" & lr & ",MATCH(SMSv1!C2,SMSv2!$K$2:$K$" & lr & ",0),3)"
    
    Dim lr1 As Long
    lr1 = Range("A2").CurrentRegion.Rows.Count
    
    Selection.AutoFill Destination:=Range("N2:N" & lr1)
    Range("N2:N" & lr1).Select
    
    'Give column name as count
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Count"
    
    Range("O2").Select
    
    'Select the last row containing data and replace 26768
    ActiveCell.Formula = _
        "=INDEX(SMSv2!$A$2:$K$" & lr & ",MATCH(SMSv1!C2,SMSv2!$K$2:$K$" & lr & ",0),7)"
        '"=INDEX(SMSv2!rng,MATCH(SMSv1!RC[-12],SMSv2!rng,0),8)"
    Range("O2").Select
    
    'Select the last row containing data and replace 25633
    Selection.AutoFill Destination:=Range("O2:O" & lr1)
    Range("O2:O" & lr1).Select
    
    'Filter by Status note: Ask about other status if present
    Range("L1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$O$" & lr1).AutoFilter Field:=12, Criteria1:= _
        "PROCESSING"
        
    'Select cell below A1
    'Range("A15398").Select
    Range("A1").Select
    ActiveCell.Offset(1, 0).Select
    'Select range to delete
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    'Delete the selection range
    Selection.EntireRow.Delete
    Selection.AutoFilter
    
    'Delete rows
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    
    'Get text to column
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("H2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 9), Array(3, 9)), TrailingMinusNumbers:=True
    
    'Name the column as Data
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Date"
    
    
    'Create pivot table
    'After removing status = Processing there will be less rows
    'Get range in worksheet SMSv1
    
    Dim rngA As Range
    Dim rngDataSMSV1 As Range
    
    Set rngA = ActiveSheet.Range("A1")
    Set rngDataSMSV1 = Range(rngA, rngA.End(xlToRight))
    Set rngDataSMSV1 = Range(rngDataSMSV1, rngDataSMSV1.End(xlDown))
    
    
    'Add new sheet and named
    Sheets.Add(After:=ActiveSheet).Name = "Virakesari_SMS_Outbound_Total"
    
    Dim PTCache As PivotCache
    Dim PT As PivotTable
   
   'Create the Cache
    Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=rngDataSMSV1)
    'Select the destination sheet
    Sheets("Virakesari_SMS_Outbound_Total").Select
    Cells(3, 1).Select
    
    'Create the Pivot table
    Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, _
        TableDestination:=Range("A3"), TableName:="Pivot1")
    
    ActiveWorkbook.ShowPivotTableFieldList = True

    With ActiveSheet.PivotTables("Pivot1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Pivot1").PivotFields("MSISDN")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveSheet.PivotTables("Pivot1").AddDataField ActiveSheet.PivotTables( _
        "Pivot1").PivotFields("Count"), "Sum of Count", xlSum
    
    'Collapse the detailed view
    Dim lastRow As Integer
    Range("A4").Select

    Dim lrSH1 As Long
    lrSH1 = Range("A4").CurrentRegion.Rows.Count
    
    ActiveCell.End(xlDown).Select
    lrSH1 = lrSH1 - 1
    Range("A4:A" & lrSH1).Select
    Selection.ShowDetail = False
    
    Sheets("SMSv1").Select
    Range("A:H").Select
    Selection.Copy
    Range("A1").Select
    
    'Paste the data as values
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Remove worksheet
    Sheets("SMSv2").Select
    ActiveWindow.SelectedSheets.Delete
End Sub

Sub VirakesiriOutboundDaily20200205V2()
'
' RupavahiniOutboundDaily20200205V2 Macro
'

'
    
    Sheets.Add(After:=ActiveSheet).Name = "Virakesari_Outbound_Daily"
    Sheets("SMSv1").Select
    Columns("H:H").Select
    Selection.Copy
    Sheets("Virakesari_Outbound_Daily").Select
    ActiveSheet.Paste
    Sheets("SMSv1").Select
    Columns("D:D").Select
    Selection.Copy
    Sheets("Virakesari_Outbound_Daily").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("SMSv1").Select
    Columns("G:G").Select
    Selection.Copy
    Sheets("Virakesari_Outbound_Daily").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Message_Count"

End Sub
Sub VirakesiriPaymentsDaily20200205V2()
'
' RupavahiniPaymentsDaily20200205V1 Macro
'

'
    Sheets.Add(After:=ActiveSheet).Name = "Virakesari_Payments_Daily"
    Sheets("Payment").Select
    Columns("G:G").Select
    Selection.Copy
    Sheets("Virakesari_Payments_Daily").Select
    ActiveSheet.Paste
    Sheets("Payment").Select
    Columns("E:E").Select
    Selection.Copy
    Sheets("Virakesari_Payments_Daily").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("Payment").Select
    Columns("B:B").Select
    Selection.Copy
    Sheets("Virakesari_Payments_Daily").Select
    Range("C1").Select
    ActiveSheet.Paste
End Sub

Sub VirakesiriDailyDashBoardAndDashBoard20200205v2()
'
' RupavahiniDailyDashBoard20200205v1 Macro
'

'
    Sheets("Virakesari_Payments_Daily").Select
    'get the last row containig data
    Dim lr As Long
    lr = Range("A1").CurrentRegion.Rows.Count
    
    Range("A1:C" & lr).Select
    Selection.Copy
    Sheets.Add(After:=ActiveSheet).Name = "Daily_Dashboard"
    Range("A1").Select
    ActiveSheet.Paste
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "SMS Count"
    Range("D2").Select
    
    'go to rupavahini outbound daily and get number of rows
    Sheets("Virakesari_Outbound_Daily").Select
    Dim lr2 As Long
    lr2 = Range("A1").CurrentRegion.Rows.Count
    
    Sheets("Daily_Dashboard").Select
    ActiveCell.Formula = _
        "=IFERROR(INDEX(Virakesari_Outbound_Daily!$A$2:$C$" & lr2 & ",MATCH(Daily_Dashboard!B2,Virakesari_Outbound_Daily!$B$2:$B$" & lr2 & ",0),3),0)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lr)
    Range("D2:D" & lr).Select
    Range("D1").Font.Bold = True
'''''''''''''''''''''''''''''''''''''''''''''
    'Create Dashboard
    'Get range of daily dashboard
    
    Dim rngA As Range
    Dim rngDataDD As Range
    
    Set rngA = ActiveSheet.Range("A1")
    Set rngDataDD = Range(rngA, rngA.End(xlToRight))
    Set rngDataDD = Range(rngDataDD, rngDataDD.End(xlDown))
    
    Sheets.Add(After:=ActiveSheet).Name = "Dashboard"
      
    'ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Daily_Dashboard!R1C1:R3398C4", Version:=6).CreatePivotTable _
        TableDestination:="Sheet6!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=6
        
    'Create the Cache
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=rngDataDD)

    'Select the destination sheet
    Sheets("Dashboard").Select
    Cells(3, 1).Select

    'Create the Pivot table
    Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, _
        TableDestination:=Range("A3"), TableName:="PivotTable1")

    ActiveWorkbook.ShowPivotTableFieldList = True
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("msisdn")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("charged_amount"), "Sum of charged_amount", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("SMS Count"), "Sum of SMS Count", xlSum
    

    Dim lrD As Long
    Range("A5").Select
    lrD = Range("A5").CurrentRegion.Rows.Count
    Range("A5:A" & lrD).Select
    Selection.ShowDetail = False
    Range("A1").Select
    
    
End Sub

Sub VirakesiriOveruseSMS20200205v1()
'
' RupavahiniOveruseSMS20200205v1 Macro
'

'
    Sheets.Add(After:=ActiveSheet).Name = "Overuse_SMS"
    
    Sheets("Dashboard").Select
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    Sheets("Overuse_SMS").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    'Get no of rows containing data
    Dim lrOU As Long
    Range("A2").Select
    lrOU = Range("A2").CurrentRegion.Rows.Count
    
    'Get the overuse SMS and assign to variable
    OverUseSMS = Range("D" & lrOU).Value
    
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Overuse SMS"
    Range("D2").Select
    ActiveCell.Formula = "=IF(C2>300,C2-300,0)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lrOU)
    Range("D2:D" & lrOU).Select
    Range("D" & lrOU).Select
    ActiveCell.Formula = "=SUM(D2:D" & lrOU - 2 & ")"
    Range("1:1").Font.Bold = True
End Sub

Sub VirakesiriProcessAllInOne20200205V2()
  Application.ScreenUpdating = False
    Call VirakesiriPayment20200130PaymentV2
    Call VirakesiriSMSOutbound31012020V2
    Call VirakesiriOutboundDaily20200205V2
    Call VirakesiriPaymentsDaily20200205V2
    Call VirakesiriDailyDashBoardAndDashBoard20200205v2
    Call VirakesiriOveruseSMS20200205v1
    
    Sheets("Final_Report").Select
    
    Range("A4").Value = OverUseSMS
    Range("C4").Value = Round(Revenue, 2)
  Application.ScreenUpdating = True
End Sub


