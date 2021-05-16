# Stock Analysis Product

## Overview of Project
Steve, a financial analyst, recently acquired a new customer and is reviewing their renewable energy stock portfolio.  He recognizes a need for better diversification of their funds.

### Purpose
Steve needs a tool that can effecienctly analyse all stocks of interest so he can better inform his clients of their diversification options.  Steve's existing program is not optmized for high throughput analysis.  His code needs to be refactored for optimized computational performance to decrease run time.

## Results
**The original code: **

    Sub allstockanalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    'format the output sheet

    Worksheets("All Stocks Analysis").Activate

    'Create title
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'initialize the array of tickers

    'initialize the array
    Dim tickers(12) As String
    
    'initialize variables in list
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    'Prepare for analysis of all tickers
    
    'initialize variables for prices
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Activate the worksheet
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop through tickers
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        'loop through rows of data
        Worksheets(yearValue).Activate
            For j = 2 To RowCount
                If Cells(j, 1).Value = ticker Then

                    'increase totalVolume by the value in the current row
                    totalVolume = totalVolume + Cells(j, 8).Value

                End If

                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    startingPrice = Cells(j, 6).Value

                End If

                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    endingPrice = Cells(j, 6).Value

                End If

            Next j
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Size = 16
    Range("B4:B15").NumberFormat = "#,##0"
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub

Analyzing all stocks by year (2017 vs 2018) as a test analysis we see clearly see different annual performances between years.

Steve's original code needed the following run time to make these analyses:

**The refactored code for high throughput analysis:**

    Sub AllStocksAnalysisRefactored()
    
    'initialize variables for timmer
    Dim startTime As Single
    Dim endTime  As Single

    'input button code
    yearValue = InputBox("What year would you like to run the analysis on?")

    'start timer
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Format the title with concatentor from input button
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
       
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerindex As Long
    tickerindex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0
    
    Next i
    
       
    '2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerindex = tickerindex + 1
            
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        Worksheets("All Stocks Analysis").Activate
        
        For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    'Format colors
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'End timer and display run time prompt
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

Analyzing all stocks by year (2017 vs 2018) again with the refactored code, we see the same results - a different annual performances between years.  The refactored code is still producing the same analysis

However, the refactored code needed less run time to make these analyses:



## Summary

- Decemeber is the worst month for launching a theater kickstarter campaign with the number of successful campaigns roughly equal to the number of failed campaigns.
- May is the best month for launching a theater kickstarter campaign with the number of successful campaigns roughly 2.15x the number of failed campaigns

- Setting a financial fundraising goal lower than 1500 improves your likelihood of success for success.

- Some limitations of the data set are the non-normalcy (right skewed) of the data when filtering by financial fundraising goal amount.  This decrease our sample size for campaigns above $1500.

- Other vairables not studied are country of origin, percentage by backer, and year over year trends.
