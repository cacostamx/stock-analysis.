# Refactoring VBA Code

## Overview of Project

Steve asked us to develop VBA code to analyze green stocks performance for years 2017 and 2018 from an Excel data set.  Although this code perfectly runs for the stocks Steve has picked, he wants to analyze more stocks.

So, after finally generating the code, we refactored it so it can run faster to analyze hundreds of stocks.

## Results

The following graphs compare both runs for the original and the refactored code for both years 2017 and 2018.

| 2017 Run for original code | 2017 for refactored code |
|-------|-------|
| ![2017 run original](/Resources/Original_2017.png)  | ![2017 run refactored](/Resources/VBA_Challenge_2017.png) |

| 2018 Run for original code | 2018 for refactored code |
|-------|-------|
| ![2018 run original](/Resources/Original_2018.png)  | ![2018 run refactored](/Resources/VBA_Challenge_2018.png) |


As it can be seen, the running time for the 2017 run was reduced fom 0.9648 to 0.1875 seconds (80.5% faster), whereas for the 2018 run it was reduced from 0.90625 to 0.125 (86% faster).

Both data sets have the same amount of rows, so I would suggest that the time difference between 2017 and 2018 runs, in both original and refactored, may be that the code may have initiated some memory allocations that helps with the second run, but the difference is unimportant.

As for why the significant reduction in time for the refactored version, I would argue that it derives from three main changes:

First, creating a ticker index to use in the main loop.  Second, the creation of array variables to store and output the information for each ticker.

```
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

Then, the third important change was to just leave one *Row* loop, by getting rid of the *Tickers* loop with the arrays variable, so we could store the required information in the first pass. 


```

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
            If Cells(i, 1) = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
                tickerIndex = tickerIndex + 1  '3d Increase the tickerIndex.
            End If
    Next i
```


## Summary

In the case of the stock analysis code, evidently the objective was accomplished as the running time was reduced significatively. Refactoring the code led to reducing lines of code and the structure itself turned into a cleaner version so it can be easier reviewed.

When refactoring code we can find ourselves with advantages and disadvantages such as:

**Advantages**

    - Of course, the main advantage is that the final code will run faster and smoothly.
    - We can optimize loops or amount of coding by the use of general variables.
    - Antoher one is that we can find some bugs or logical errors that are not evident in the first run because the code may run without errors until specific cases.
    - We can make the structure cleaner and add comments so further reviews by peers can be easier.

**Disadvantages**

    - I think that the main disadvantage of refactoring is the time employed in going through all the code. Especially if it is long enough.
    - It can derive into complexity instead of just the original objective (i.e. making it cleaner, faster running, etc.).
    - One could get lost in the changes done after a lot of refactoring. Hoping comments will help with this.



Regarding our code refactoring, the main changes to the code were taking out the *Tickets* loop from the *Rows* loop, and creating array variables for the volume and prices. Also reviewing-wise, these changes create an easier way to review the code.

On the cons side I did found myself investing a lot of time because I had an error in the array dimensioning so I had to debug the code several times to find it.
