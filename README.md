# Green Stocks Analysis
## Overview of Project
The purpose of this project was to refactor a VBA script used to analyze the performance of a collection of green energy stocks. The original code was functional, but required running multiple loops through the dataset, increasing its run time. By refactoring the code to require only one loop, I improved its efficiency, allowing for the efficiency of a much larger dataset.

## Results
### Stock Performance
Overall, this group of stocks peformed significantly better in 2017 than 2018. All stocks except TERP had postive returns in 2017; only ENPH and RUN, however, continued to gain value in 2018. ENPH was the strongest performer across both years. 
![2017 performance](https://user-images.githubusercontent.com/79542537/111043830-ca1b5a00-8412-11eb-859d-2e26642bdbb9.PNG)
![2018 performance](https://user-images.githubusercontent.com/79542537/111043843-e15a4780-8412-11eb-8801-b6d6785346af.png)



### Refactoring the Script
After refactoring the script, run time improved significantly. Compare the run time of the original script vs. the refactored version.

![2017 original](https://user-images.githubusercontent.com/79542537/111043475-c71f6a00-8410-11eb-95b7-e01a15937f56.png)
![2017 msgbox](https://user-images.githubusercontent.com/79542537/111043476-cbe41e00-8410-11eb-8015-bf3c90592199.png)

The original script was slower, because it had to loop through the dataset multiple times, outputting the results for each individual stock, and then repeating the process until the full results were collected. To collect the full results for our dataset, it had to repeat the following loop 12 times.

'4)Loop through the tickers
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
      
      Next i'
      
 The refactored code instead only needs to loop through the dataset once, storing its results in arrays before outputting them to a worksheet. 

## Summary
#### What are the advantages or disadvantages of refactoring code?
Refactoring code is time consuming work. Time that could be spent on a new project instead is used reviewing and editing existing code. But the effort to refactor may be worthwhile to improve the efficiency and logical structure of your code. Refactored code may be easier for another programmer to understand and it edit, and it may make a script run faster and use less memory.

#### How do these pros and cons apply to refactoring the original VBA script?

The original VBA script used in this project was functional, produced correct results, and ran in a short amount of time. But it only needed to output results for 12 stocks, and may not have performed as well with a larger data set. By refactoring the script to use fewer loops, its run time improved significantly. The resulting script is not only faster, but can likely handle a much larger dataset, making it more expandable.
