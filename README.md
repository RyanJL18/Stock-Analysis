# Analysis of Stocks Performance over 2017 and 2018
---
## Overview
Using the data provided, we are tasked with assisting "Steve" in finding the performance of each stock. Following the initial utilization of the code, we are going to refactor the code to ensure that it is running as efficiently as possible. Efficiency is important when concerning these stocks, because in the event that more stocks are added to the data, we need to ensure that the code is running at a speed that will prevent any slow down.

# Results
---
## 2017 Performance
2017 was a very positive year in terms of performance. Every stock except for **TERP** had a positive return. Stocks that performed extremely well include: **DQ**, **ENPH**, **FSLR**, and **SEDG**; with each of these stocks having a return of *over* 100%. This can be considered a good year for investors. 
## 2018 Performance
Sadly, 2018 was a bit on the opposite side of the spectrum. Almost all stocks displayed a negative return, with the exceptions being **ENPH** and **RUN**. Specifically **ENPH** displayed positive returns in both 2017 and 2018, making it a good stock for anyone that had chosen to invest. However, Steve's parents choosing to invest in **DQ** was not the best choice as **DQ** had the highest *negative* return at 62.6%. 
## Code Performance
Code performance was improved dramatically with the initial code dropping from 1.75 seconds to 0.125 seconds when calculating the data for 2018.


![2018 VBA Run.png](https://github.com/RyanJL18/Stock-Analysis/blob/main/2018%20VBA%20Run.png)

![2018 VBA Refactored Run.png](https://github.com/RyanJL18/Stock-Analysis/blob/main/2018%20VBA%20Refactored%20Run.png)


Calculating the date for 2017, you can see a drop from 1.67 seconds to just over a tenth of a second. 


![2017 VBA Run.png](https://github.com/RyanJL18/Stock-Analysis/blob/main/2017%20VBA%20Run.png)

![2017 VBA Refactored Run.png](https://github.com/RyanJL18/Stock-Analysis/blob/main/2017%20VBA%20Refactored%20Run.png)

This leads us to determine that the refactored code below is running much faster than the initial code, with speed increasing by roughly 92%. 

# Refactored Code

    '1a) Create a ticker Index

    tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'If the next row's ticker is different, increase tickerIndex.
        For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
 
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

    '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
        
        
        End If

    Next i

     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

# Summary
## The Duality of Refactoring
Refactoring code has advantages and disadvantages. A perfect example of an advantage is one that we can see in action with the information above; it runs faster, much faster. This is great news! We want to save as much time as possible, especially as data sets become larger and more numbers are involved. However, a disadvantage of refactoring is that it can be hard to read for someone who is not familiar with your code. It would make sense for files and code to be shared in an office setting where it is needed to be viewed by multiple people. Fresh eyes looking at a condensed version of you code may be confused and require an explaination for something that may have been easily noticable in the original code before refactoring.
___
In this specific instance, refactoring helped us by bringing the time down drastically. While it is a fraction of a second in this project, that time saved will only grow as the data, number of tickers, or stocks increase. The disadvantage of refactoring in this instance is that it can be time consuming. With an increase in time that is so slight going from 1 second to a fraction of a second, more time would have absolutely been saved by forgoing the refactoring process and sticking with the original code. Going over lines of code multiple times can be tricky and time consuming. There were many times that I ran into errors that required me to look through each line of code to find an error, with the most common culprit being that I had accidentally keyed in an extra letter on any given line.
