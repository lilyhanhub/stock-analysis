# Stock Analysis with Excel VBA

## Overview of Project
The purpose of this project is to refactor a Microsoft Excel VBA code for Steve to analyze the performance of certain stocks in the years 2017-2018, and to provide his parents with some insight if the stocks are worth investing. There was a similar code completed earlier. Now the goal for this round is to increase the efficiency of the original code.

The raw dataset comes with two worksheets of 12 stocks spanning two consecutive years. Each stock information consists of a ticker value, date, daily opening/closing/adjusted closing price, the highest/lowest price, and the volume of daily transactons. The goal is to create a new worksheet called "All Stocks Analysis"  and to find the total daily volume and yearly return rate for each stock by looping over all of the tickers.
For more information about the dataset: [VBA_Challenge](/VBA_Challenge.xlsm)

## Results
Below are the snippets of the refactored code.   
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

### Comparison of stock performance between 2017 and 2018
For the twelve stocks Steve selected for this analysis, there's a much higher average return in 2017, +67.5% compared to -8.5% in 2018. There were two stocks that performed the best in both years: ENPH and RUN, followed by two additional stocks with overall more positive returns: SEDG and VSLR. 
### Comparison of execution time between the original script and the refactored script
The refactored code successfully made the VBA scripts run much faster. The execution time significantly decreased from 0.53s to 0.09s (2017) and from 0.54s to 0.10s (2018). 

#### Refactored results for 2017
![Refactored 2017](/Resources/2017all_pics.png)
#### Original results for 2017
![Original 2017](/Resources/2017all_pics_original.png)

#### Refactored results for 2018
![Refactored 2018](/Resources/2018all_pics.png)
#### Original results for 2018
![Original 2018](/Resources/2018all_pics_original.png)
## Summary 
1. What are the advantages or disadvantages of refactoring code?
 * There are clear advantages of refactoring code. It not only saves the processing time, but also makes the code more efficient, functional, structured, and easier for the future users to read. However, the refactoring process may take much time to figure out ways to simplify the original code, especially for a novice programmar. Also it's risky to refactor code when the application is big and the existing code does not have proper test cases.

2.  How do these pros and cons apply to refactoring the original VBA script?
 * Based on the above results, refactoring the original VBA script proved worth of the efforts with about 82% less execution time. Also, the refactored
   stock analysis code offers us the framework to analyze larger datasets with more stocks and a longer timeline for comparision. However, refactoring in VBA is quite challenging and there was a lot of troubleshooting and testing pieces of code.
