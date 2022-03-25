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

# Refactored Code

     tickerIndex = 0


    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single


    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i

    
    For i = 2 To RowCount

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

  
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
  
    tickerIndex = tickerIndex + 1
    End If

    Next i
  

    For i = 0 To 11
    
      Worksheets("All Stock Analysis").Activate
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

# Summary



---
**Fig1:**
![Formula Average.png](https://github.com/RyanJL18/Kickstarter-Analysis/blob/main/Formula%20Average.png)
**Fig2:**
![Formula Epoch Time.png](https://github.com/RyanJL18/Kickstarter-Analysis/blob/main/Formula%20Epoch%20Time.png)
