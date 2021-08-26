# Stock-analysis

## Overview

The purpose of the project was to collect selected stock information for 2017 and 2018 to determine if the stocks are worth investing. Two variations of the code were written to determine which method was most efficient in collecting the desired stock information.

## Results

The updates I made to the code changed the way the code ran through the 12 ticker symbols. Instead of running the code multiple times for each ticker to get the final results, the refactored code ran through the tickers only once, which is the reason for the decreased time. Below is the refactored code and the screenshots showing the running time it took to run for both 2017 and 2018. 

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
     Dim tickerVolumes(12) As Long
     Dim tickerStartingPrices(12) As Single
     Dim tickerEndingPrices(12) As Single
     
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
     For i = 0 To 11
       tickerVolumes(i) = 0
       tickerStartingPrices(i) = 0
       tickerEndingPrices(i) = 0
    
    Next i
     
    ''2b) Loop over all the rows in the spreadsheet.
     For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                        
             End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            

            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
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


![VBA_Challenge_2017](https://user-images.githubusercontent.com/88639467/131036278-14d0f958-01d9-4951-9f7f-9a1b5beea560.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/88639467/131036294-c5ddbf4a-2249-4ca3-8dc7-7173a7d2d3c4.png)

## Summary

### Advantages of Refactoring Stock Analysis

The most significant different between the orginal and refactored code is the decrease in running time. The original analysis took almost one and half seconds to run, whereas the refactored code took approximately 0.28 seconds to run.

### Pros & Cons of Refactoring

Refactoring helps make the code cleaner, more organized and more efficient by taking fewer steps and using less memory. The added benefit of refactoring is to make the code easier to read for future users who either need to use the code or take over a project. Unfortunately, refactoring may not always be possible due to certain contraints. These include having applications that are too large or not having the proper test cases for the existing codes, which may pose risk if the code were to be refactored. 
