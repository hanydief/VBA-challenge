Attribute VB_Name = "Module1"
Sub YearlyStockData()

' Applying changes to all active sheets ********Bonus********
For Each ws In Worksheets       'run for all sheets on file
Dim WorksheetName As String

WorksheetName = ws.Name

'declaration
Dim rowcount As Long
Dim tickercount As Long
Dim tickerchange As Long

'Column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Change %"
    ws.Range("L1").Value = "Total Stock Volume"
'Summary Data headings
    ws.Range("N2").Value = "Greatest Increase %"
    ws.Range("N3").Value = "Greatest Decrease %"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
MsgBox ("Processing All Sheets, same message will popup once for each sheet, Please Hang tight , a pop up message Will let you know once done")     'message to let user know it is finished since process takes too long
 
rowcount = Cells(Rows.Count, 1).End(xlUp).Row   'assigns rowcount to expression
tickercount = 2                                 'assigns Starting tickercount
customeindex = 2                                'assigns starting CustomeIndex

'Looping through all active sheets in the woorksheet
ws.Cells(customeindex, 1).Value = ws.Cells(2, 1).Value

For i = 2 To rowcount
 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then          'if current Ticker cell doesn't equal next Cell do calculations otherwise increase i by 1 using the for loop
        ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value   ' returns ticker value
        ws.Cells(tickercount, 10).Value = "$" & Round(ws.Cells(i, 6).Value - ws.Cells(customeindex, 3).Value, 2)                        ' returns yearly Change value
        ws.Cells(tickercount, 11).Value = ((ws.Cells(i, 6).Value / ws.Cells(customeindex, 3).Value) - 1) * 100 & "%"                    ' returns yearly % Change value
        ws.Cells(tickercount, 12).Value = ws.Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i, 7), ws.Cells(customeindex, 7)))     'Total volume

        
    'Conditional color formatting
        CondColor = ws.Cells(tickercount, 10).Value
            Select Case CondColor
                Case Is > 0
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                Case Else
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 0
            End Select
        
        tickercount = tickercount + 1      'adds up 1 to the ticker count for every new ticker  code
        customeindex = i + 1               'adds up 1 to the custome index for every time ticker code changes

        End If
Next i

'*************Bonus**************
'Functionality Data Greatest Increase/Decrease % and Volume: Tickers & Values
    'Summary Data Greatest Increase %
    ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100 & "%"
    IncreaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
    ws.Range("O2") = ws.Cells(IncreaseNumber + 1, 9)
    'Summary Data Greatest Decrease %
    ws.Range("P3").Value = WorksheetFunction.Min(Range("K2:K" & rowcount)) * 100 & "%"
    DecreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
    ws.Range("O3") = ws.Cells(IncreaseNumber + 1, 9)
    'Summary Data Total Volumes
    ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
    IncreaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)
    ws.Range("O4") = ws.Cells(IncreaseNumber + 1, 9)

'Auto adjust the width of columns
ws.Columns("J:Q").AutoFit
ws.Rows(1).Font.Bold = True
ws.Columns("J:J").NumberFormat = "$[Black]-0.0#"                                                                                ' Makes sure negative numbers doesn't show in red since cell color might be red
Next ws

MsgBox ("Finished All Sheets")  'message to let user know it is finished since process takes too long

End Sub
'Runs for a single sheet Only
Sub ThisYearStockData()
'declaration
Dim rowcount As Long
Dim tickercount As Long
        
'Column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Change %"
    Range("L1").Value = "Total Stock Volume"
'Summary Data headings
    Range("N2").Value = "Greatest Increase %"
    Range("N3").Value = "Greatest Decrease %"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
 
 MsgBox ("Processing Current Sheet Only, Please Hang tight it would take few seconds, a pop up message Will let you know once done")     'message to let user know it is finished since process takes too long
 
rowcount = Cells(Rows.Count, 1).End(xlUp).Row       'assigns rowcount to expression
tickercount = 2                                     'assigns tickercount
customeindex = 2                                     'assigns custome index
'Looping thru rows
For i = 2 To rowcount
    If Cells(i, 1).Value <> Cells(i + 1, 1) Then            'if current Ticker cell doesn't equal next Cell do calculations otherwise increase i by 1 using the for loop
        Cells(tickercount, 9).Value = Cells(i, 1).Value     ' returns ticker value
        Cells(tickercount, 10).Value = "$" & Round(Cells(i, 6).Value - Cells(customeindex, 3).Value, 2)                 ' returns yearly Change value
        Cells(tickercount, 11).Value = ((Cells(i, 6).Value / Cells(customeindex, 3).Value) - 1) * 100 & "%"             ' returns yearly % Change value
        Cells(tickercount, 12).Value = Application.WorksheetFunction.Sum(Range(Cells(i, 7), Cells(customeindex, 7)))    'Total volume

        'Conditional color formatting
        
        CondColor = Cells(tickercount, 10).Value
            Select Case CondColor
                Case Is > 0
                    Cells(tickercount, 10).Interior.ColorIndex = 4
                Case Is < 0
                    Cells(tickercount, 10).Interior.ColorIndex = 3
                Case Else
                    Cells(tickercount, 10).Interior.ColorIndex = 0
            End Select
        tickercount = tickercount + 1      'adds up 1 to the ticker count for every new ticker  code
        customeindex = i + 1               'adds up 1 to the custome index for every time ticker code changes
    End If
Next i

'*************Bonus**************
'Functionality Data Greatest Increase/Decrease % and Volume: Tickers & Values
    'Summary Data Greatest Increase %
    Range("P2").Value = WorksheetFunction.Max(Range("K2:K" & rowcount)) * 100 & "%"
    IncreaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    Range("O2") = Cells(IncreaseNumber + 1, 9)
    'Summary Data Greatest Decrease %
    Range("P3").Value = WorksheetFunction.Min(Range("K2:K" & rowcount)) * 100 & "%"
    DecreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    Range("O3") = Cells(IncreaseNumber + 1, 9)
    'Summary Data Total Volumes
    Range("P4").Value = WorksheetFunction.Max(Range("L2:L" & rowcount))
    IncreaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowcount)), Range("L2:L" & rowcount), 0)
    Range("O4") = Cells(IncreaseNumber + 1, 9)

'Auto adjust the width of the columns & color all fonts to black
Columns("J:Q").AutoFit
Rows(1).Font.Bold = True
Columns("J:J").NumberFormat = "$[Black]-0.0#"                                                                   ' Makes sure negative numbers doesn't show in red since cell color might be red

MsgBox ("Finished this Sheet")  'message to let user know it is finished since process takes too long

End Sub

'Clear CURRENT active sheet ONLY
Sub ClearThisSheet()

' ClearAll Macro
    Range("H:P").Clear
    Range("H:P").ClearFormats
    
End Sub
    
'Clear All active sheets
Sub ClearALLSheets()

For Each ws In Worksheets  'Applying clear to all active sheets
Dim WorksheetName As String

WorksheetName = ws.Name

' ClearAll Macro
    ws.Range("H:P").Clear
    ws.Range("H:P").ClearFormats
    
Next ws
End Sub
