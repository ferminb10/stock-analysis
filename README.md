# stock-analysis
Performing analysis on green stock data to uncover trends.

## Overview of Project
This weeks challenge helps us build upon our skills learned in the VBA module. The client wants to do a do research for his parents. He'd like to possibly expand the dataset to the entire stock market over the last few years.
## Results
The analysis had two part task. First, a solution code was refractored to loop through all stock data one time in order to collect the same information that you did in this module. Then, it was determined if that our refactored code successfully performed better than the original script, the results shown below.

### 2017 Refractored 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107658895/175761138-60faaf6e-37bd-4d96-a29c-ecc37ee3e70f.png)
![VBA_Challenge_2017_timer](https://user-images.githubusercontent.com/107658895/175760314-8ddae46c-3025-489d-9b83-978c878992f3.png)
### 2018 Refractored
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107658895/175760753-bb322daa-9686-4cdf-9c32-312b4ffdefb8.png)
![VBA_Challenge_2018_timer](https://user-images.githubusercontent.com/107658895/175760756-14761ff9-d98c-4fbd-bc73-746dc903f7f1.png)

### Original 2017 & 2018
![Original](https://user-images.githubusercontent.com/107658895/175761234-a23bb236-69ac-4d8e-83fe-fe8925c1b97e.png)
# Code:

    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim TickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single
    Dim TickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    TickerVolumes(i) = 0
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        TickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
    End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        TickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    End If


            '3d Increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
            
        'End If
    End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
     Cells(4 + i, 1).Value = tickers(i)
     Cells(4 + i, 2).Value = TickerVolumes(i)
     Cells(4 + i, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

## Summary
Based on the results of this test the refractered code was a more efficient version of the original script. This is one of the benefits of refractoring code. You can take code and try to optimize. The limitations of this dataset is that it only provides a small dataset. This script doesn't do a good job of testing its limits. I was curious to know how long it would take to run the S&P 500. One of the disadvantages of refractoring code for the particular application of analyzing thousands of stocks is that the program may take a long time to execute. You can easily make this to your advantage if you're well educated in economics and have it tailored to a specific set that you want. The refractored script proved to be a much more efficient than the original based on the executed times. The faster it can be the more applicible it can be for a wider range of datasets.
