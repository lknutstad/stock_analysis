# Stock Analysis Program Refactoring

## Overview of Project
Steve, a financial analyst, recently acquired a new customer and is reviewing their renewable energy stock portfolio.  He recognizes a need for better diversification of their funds.

### Purpose
Steve needs a tool that can effecienctly analyse all stocks of interest so he can better inform his clients of their diversification options.  Steve's existing program is not optmized for high throughput analysis.  His code needs to be refactored for optimized computational performance to decrease run time.

## Results
## The original code:

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

- To test steve's original program, an analysis of all stocks by year (2017 vs 2018) is used. A clear difference in annual performance between years is observed. Steve's original code needed approximate 0.55 seconds to run these analyses.

![image](https://user-images.githubusercontent.com/16930677/118395531-c321eb00-b5ff-11eb-9d6e-03e9a3aba6ba.png)
![image](https://user-images.githubusercontent.com/16930677/118395523-b8ffec80-b5ff-11eb-89fe-12eb47476f7e.png)





## The refactored code:

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

- Analyzing all stocks by year again with the refactored code, we see the same results - a different annual performances between years.  The refactored code is still producing the same analysis. However, the refactored code needed only 0.1 seconds (5x less than original) run time to make these analyses:
![image](https://user-images.githubusercontent.com/16930677/118395547-dd5bc900-b5ff-11eb-9e14-06d12e44db03.png)






## Summary

Refactoring code is an important step in developing a robust program for analysis and this project for Steve demonstrates that.  After refactoring, the original program was 5 times faster due to a combination of several factors - fewer lines of code, less memory used, etc.  

One major area for refactoring Steve's orginal program was in the "for" loop.  The original code contained several steps inside the "for" loop where as the reafactored code broke that large loop down into several smaller loops and pre loop steps.  With the program running 5 times faster, this allows Steve to run 5 times more data!


