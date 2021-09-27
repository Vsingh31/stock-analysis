# stock-analysis
## Overview of Project
purpose of this project is refactoring your old code and masure performance.To the determind  whether refactor VBA code, successfully made the VBA script run faster by measuring  performance. 
### Results
* For refactoring code and measuring performance, I Created a tickerIndex variable and set it equal to zero before iterating over all the rows.
  tickerIndex = 0
* Created three output array.In VBA array starts with 0.We anitialize array with **DIM** Keyword.there was one more array named tickers() that created for holding all the        tickers.
   - Dim tickerVolumes(0 To 11) As Long
    -Dim tickerStartingPrices(0 To 11) As Single
    -Dim tickerEndingPrices(0 To 11) As Single
