# Green Stocks Analysis
## Overview of Project
The purpose of this project was to refactor a VBA script used to analyze the performance of a collection of green energy stocks. The stock data and VBA scripts can be found in the workbook, [VBA_Challenge](VBA_Challenge.xlsm). The original script, AllStocksAnalysis, requires running multiple loops through the dataset, increasing its run time. The refactored script, AllStocksAnalysisRefactored, only needs to loop through the data once. This improves its efficiency and will allow for the analysis of a much larger dataset in the future.

## Results
### Stock Performance
Overall, this group of stocks performed significantly better in 2017 than 2018. All stocks except TERP had positive returns in 2017; only ENPH and RUN, however, continued to gain value in 2018. ENPH was the strongest perfomer across both years.

![2017 performance](https://user-images.githubusercontent.com/79542537/111043830-ca1b5a00-8412-11eb-859d-2e26642bdbb9.PNG)

![2018 performance](https://user-images.githubusercontent.com/79542537/111043843-e15a4780-8412-11eb-8801-b6d6785346af.png)

### Refactoring the Script
After refactoring the script, run time improved significantly. Compare the run time of the original script vs. the refactored version.

![2017 original](https://user-images.githubusercontent.com/79542537/111043475-c71f6a00-8410-11eb-95b7-e01a15937f56.png)
![2017 msgbox](https://user-images.githubusercontent.com/79542537/111043476-cbe41e00-8410-11eb-8015-bf3c90592199.png)

The original script is slower because it relies on a nested loop that runs through the entire dataset once for each stock. The following code outputs its results to a worksheet and then repeats itself until the analysis is complete. For our dataset, the nested loop needs to run 12 times.

```
'4a)Loop through the tickers
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
        '5) Loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
            '5a) Get total volume
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
      
      '6) Output data for current ticker
      Worksheets("All Stocks Analysis").Activate
      Cells(4 + i, 1).Value = ticker
      Cells(4 + i, 2).Value = totalVolume
      Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
      
      Next i
   ```  
      
The refactored code instead only needs to loop through the dataset once, storing the values for each stock in an array before outputting the final results. The script uses these output arrays to store the stock values.
```
    Dim tickerVolumes(12) As String
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
This allows to collect a full set of results in a single loop, completing the analysis much more quickly.
```
For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
  
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
        Next i
 ```
## Summary
#### What are the advantages or disadvantages of refactoring code?
Refactoring code is time consuming work. Time that could be spent on a new project instead is used reviewing and editing existing code. But the effort to refactor may be worthwhile to improve the efficiency and logical structure of your code. Refactored code may be easier for another programmer to understand and it edit, and it may make a script run faster and use less memory.

#### How do these pros and cons apply to refactoring the original VBA script?

The original VBA script used in this project was functional, produced correct results, and ran in a short amount of time. But it only needed to output results for 12 stocks and may not have performed as well with a larger data set. By refactoring the script to use fewer loops, its run time improved significantly. The resulting script is faster and should be able to process a much larger dataset.
