# Stock Analysis with Excel VBA

## Overview of Project
The purpose of this project is to refactor a Microsoft Excel VBA code for Steve to analyze the performance of certain stocks in the years 2017-2018, and to provide his parents with some insight if the stocks are worth investing. There was a similar code completed earlier. Now the goal for this round is to increase the efficiency of the original code.

The raw dataset comes with two worksheets of 12 stocks spanning two consecutive years. Each stock information consists of a ticker value, date, daily opening/closing/adjusted closing price, the highest/lowest price, and the volume of daily transactons. The goal is to create a new worksheet called "All Stocks Analysis"  and to find the total daily volume and yearly return rate for each stock by looping over all of the tickers.
For more information about the dataset: [VBA_Challenge](/VBA_Challenge.xlsm)

## Results
> Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. The analysis is well described with screenshots and code
```
 '1a) Create a ticker Index
        tickerIndex = 0

 '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
       
    
 ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For j = 0 To 11
        tickerVolumes(tickerIndex) = 0
        Next j
        
 ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
  ```
Refactored results for 2017
![Refactored 2017](/Resources/2017all_pics.png)
Original script results for 2017
![Original 2017](/Resources/2017all_pics_original.png)
Refactored results for 2018
![Refactored 2018](/Resources/2018all_pics.png)
Original script results for 2018
![Original 2018](/Resources/2018all_pics_original.png)
## Summary 
1. What are the advantages or disadvantages of refactoring code?


2. How do these pros and cons apply to refactoring the original VBA script?
