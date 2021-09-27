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
* Created a outer For loop with iterator tickerIndex that goes through all of the ticker inside the for loop get the tickerindex from the tickers() array and initialize the tickerVolumes to zero also.
  - For tickerIndex = 0 To 11
  - tickerVolumes(tickerIndex) = 0
* Created a inner for loop with j iterator and increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.And       Used the tickerIndex variable as the index.
  - For j = 2 To RowCount 
  - If Cells(j, 1).Value = tickers(tickerIndex) Then
  - tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
  - End If
* Using Multiple condition check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices   variable.
  - If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
  - tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
  - End If
* Similarly Multiple condition check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
  - If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
  - tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
  - End If
* Looping through your arrays to output the Ticker, Total Daily Volume, and Return in Output worksheet.
  - Worksheets("All Stocks Analysisrefactored").Activate
  - Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
  - Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
  - Cells(4 + tickerIndex, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
  
 * **My Outpur Worksheet for (All Stocks-2017) looks like following:**
 
![Data_Resultof2017](https://user-images.githubusercontent.com/90277142/134859367-e31a931c-310d-4e87-9282-ee7cf33afc76.png)

