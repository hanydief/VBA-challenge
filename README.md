# VBA-challenge
 Module2_VBA-challenge

 # Table of Content
 This Readme file has:
 * Assignment Explanation - Starts Line 11
 * the Assignmnet Module 2 VBA CHallenge requirments - Starts Line 30
 * Then starting Line - Starts Line 110

*********************************************************************************************************************************************
 # Assignment Explanation
* Worksheet has 3 tabs all are working with outstanding results
* Each Sheet contains 4 buttons:
    1. Can run all sheets
    2. Clear & reset all Sheets
    3. Run Current active sheet individually Only
    4. Clear & reset Current Active sheet only
* When you run the sheet or sheets you will get a popup message notifying you process is running, each sheet has a single popup message & one single popup message at the end notifying its done to avoide any confussion since running file takes few seconds
* Since Yearly Change column will have Negative (-ve) values & depends on excel column format some times (-ve)s might show in red & since Assignment requires negative to she RED cells number will not be visable therefore this column has color formating at the end to switch RED to BLACK.
* Outputs are:
    1. The ticker symbol
    2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    4. The total stock volume of the stock. matching requested image.
    5. BONUS: Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the requested image.
    6. BONUS: Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.


********************************************************************************************************************************************
# VBA Homework: The VBA of Wall Street

## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment, you will use VBA scripting to analyze generated stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

### Before You Begin

1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**.

2. Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock Market Analyst

![alt=""](Images/stockmarket.jpg)

## Instructions

Create a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

The result should match the following image:

![moderate_solution](Images/moderate_solution.png)

## Bonus

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

![hard_solution](Images/hard_solution.png)

Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.

## Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3 to 5 minutes.

* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button.

* Some assignments, like this one, contain a bonus. It is possible to achieve proficiency for this assignment without completing the bonus. The bonus is an opportunity to further develop your skills and be rewarded extra points for doing so.

## Submission

To submit, please upload the following to GitHub:

  * A screen shot for each year of your results on the multi-year stock data.

  * VBA scripts as separate files.

Be sure to commit regularly to your repository and that it contains a README.md file.

After saving your work, create a shareable link and submit the link to <https://bootcampspot-v2.com/>.

## Rubric

[Unit 2 Rubric - VBA Homework - The VBA of Wall Street](https://docs.google.com/document/d/1OjDM3nyioVQ6nJkqeYlUK7SxQ3WZQvvV3T9MHCbnoWk/edit?usp=sharing)

## References

* Dataset generated by Trilogy Education Services, LLC.

© 2022 Trilogy Education Services, a 2U, Inc. brand. All Rights Reserved.

***************************************************************************************************************************************

# Assignment Code copy



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
