# Stock Analysis using VBA

# Overivew 
Write a macro that can loop through spreadsheets of stock market data and display the name, volume, and return percentage. Macro utilizes user input and a timer to track how long the caculations take.

# Results
The best investments are ENPH and RUN as they had a positive gain over a two year spread, with the worst investment being TERP which suffered losses both years. 
The original run time for the 2017 analysis was 0.66 seconds and for 2018 it was 0.72 seconds.
![orignial2017](https://user-images.githubusercontent.com/88058739/128825652-08992220-2107-4eb1-b52e-1606fb811780.png)
![orignial2018](https://user-images.githubusercontent.com/88058739/128825665-e8a1bdb7-fa61-4c59-be70-6630d5e523bd.png)

The refactored code dropped the 2018 run time to 0.69 but slightly increased on the 2017 run at 0.67 seconds. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/88058739/128826030-788bd6b0-495c-4a6d-92d8-4dbaf178a4c9.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/88058739/128826032-81bcbc62-2473-483f-97cc-6760652d4295.png)

The main difference between the two is the use of arrays to hold larger amounts of information in the refactored code vs recalculating a varible with each loop in the original. The refactored code also uses multiple seperate loops and shorter IF statements. 
Example of arrays used and the loop used to collect data from the refactored code: 
```
`1a) Create a ticker Index
    tickerIndex = 0
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
`2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
        tickerVolumes(i) = 0
Next i

   Worksheets(yearValue).Activate
`2b) Loop over all the rows in the spreadsheet.
For j = 2 To RowCount
        '3a) Increase volume for current ticker
    If Cells(j, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        'set starting price
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    End If
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        'set ending price
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
    End If
Next j
```
tickers in this code is a String array storing the names of the stocks.
# Summary 
## What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code could be optimizing run time, increasing memory, and increasing the amount of data the code can process. Disdvantages could be "breaking" the code, or creating a code that's longer and doesn't increase efficiency to make it worthwhile. 
## How do these pros and cons apply to refactoring the original VBA script?
In the process of refactoring this code, I ran into challenges with getting the ticker prices to store in the array, and I feel like the time spent fixing the code vs the decreased run time was not constructive. However, given my relative newness to this profession, it could be that I simply need more practice. 
