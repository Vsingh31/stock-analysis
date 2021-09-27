# stock-analysis
## Overview of Project
purpose of this project is refactoring your old code and masure performance.To determind  whether refactor VBA code, successfully made the VBA script run faster by measuring  performance. 
### Results
* Created new subroutine with name **"AllStocksAnalysisrefactor"**
* Created a output worksheet "All Stocks Analysisrefactored" in Excel and activate it in code in multiple times. so all we can interact with that worksheet.and where we will put the output of the stocks analysis.
* 
* I formatted my output worksheet with putting "All Stocks (" + yearValue + ")" string in A1 cell to make a Title
* Range("A1").Value = "All Stocks (" + yearValue + ")"
* Then i created three columns in output worksheet with headers-Ticker,Total Daily Volume and Return.Assign these strings in third row and first,second and third columns(A3,B3,C3)
   - Cells(3, 1).Value = "Ticker"
   - Cells(3, 2).Value = "Total Daily Volume"
   - Cells(3, 3).Value = "Return"
* Created three output array.In VBA array starts with 0.We anitialize array with **DIM** Keyword.there was one more array named tickers() that created for holding all the           tickers.
   - Dim tickerVolumes(0 To 11) As Long
   - Dim tickerStartingPrices(0 To 11) As Single
   - Dim tickerEndingPrices(0 To 11) As Single
* For refactoring code and measuring performance, I Created a tickerIndex variable and set it equal to zero before iterating over all the rows.
  tickerIndex = 0
* Found the no. of rows with the help of this code to loop over.
   - RowCount = Cells(Rows.Count, "A").End(xlUp).Row  
* Created a outer For loop that goes through all of the ticker inside the for loop get the tickerindex from the tickers() array and initialize the tickerVolumes to zero.
  - For tickerIndex = 0 To 11
  - tickerVolumes(tickerIndex) = 0
* Created a inner for loop with j iterator and increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
  - For j = 2 To RowCount 
           - If Cells(j, 1).Value = tickers(tickerIndex) Then
               - tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
          - End If

