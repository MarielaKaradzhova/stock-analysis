
#### Stock Analysis

### Overview of Project

The purpose of this project was to expand the dataset provided by our client, Steve. With the use of VBA we refactored code in order to analyze many stocks in multiple years using automated commands.
In this project, we focused on 12 stocks in particular, with interest in comparing all stocks available to the DQ stock. We also extracted the total daily volume and return of each stock in the array. 

### Results

Below is a link to the Excel workbook which contains the relevant code for this project:
https://github.com/MarielaKaradzhova/stock-analysis/blob/main/VBA_Challenge.xlsm

Included below is the section of relevant code which has been refractored:

    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        'End If
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
      
        'If  Then
        
        End If
        
        '3d Increase the tickerIndex.
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        
        'End If
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
To determine the preformance of all 12 stocks for 2017 and 2018 we created code which contained arrays and loops and also  used  "IF"statements :

2017                                                                    |2018
:----------------------------------------------------------------------:|:------------------------------------------------------------------------------:
![](https://github.com/MarielaKaradzhova/stock-analysis/blob/main/All_Stocks_2017.png) | ![](https://github.com/MarielaKaradzhova/stock-analysis/blob/main/All_Stocks_2018.png)

From the tables above, we can see that in 2017 all stocks but TERP was growing in value, where DQ was the stock with highest return at almost %200. 
In 2018, all stocks besides RUN and ENPH had a positive return. Additionally, it is important to note that RUN and ENPH were the only two stocks which displayed a positive return in 2017 and 2018, where ENPH had the largest gain of all stocks for both years.
It would be interesting and possibly beneficial to investigate the return of RUN and ENPH in previous years to get an better idea of their record of growth and compare these two stocks to DQ in order to deepen our analysis.

### Summary

##  Advantages and Disadvantages of Refactoring Code

The use of refactoring provides a framework of code to use for a given analysis. in the case of this project, it allowed us to skip writing some basic code which is the framework of our analysis. Refractoring code helped us to skip the basics and focus on creating code in less time, which is as concise and clean as possible. 
One disadvantage of refractored code is that it isn't code that we wrote ourselves. If the code is not properly explained with comments we might quickly get lost trying to apply it to our current analysis. In turn, that could lead to mistakes and problems that we will then need to fix ourselves, so the code may not be as efficient as we would of thought initially. 

Specific to Stock Analysis code, the biggest benefit of refractoring the code was the improvement in run time. Below are the timer results:

2017                                                                     |2018
:-----------------------------------------------------------------------:|:----------------------------------------------------------------------------:
![] (https://github.com/MarielaKaradzhova/stock-analysis/blob/main/VBA_Challenge_2017.png)| ![](https://github.com/MarielaKaradzhova/stock-analysis/blob/main/VBA_Challenge_2018.png)
