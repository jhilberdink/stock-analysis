# Green Stocks Analysis
## Overview of Project
The purpose of the project was to compare the performance of a collection of green energy stocks from 2017 to 2018. The analysis was conducted in Excel using VBA. The original code was refactored to improve its speed and efficiency.

## Results
### Stock Performance
Overall, this group of stocks peformed significantly better in 2017 than 2018. All stocks except TERP had postive returns in 2017; only ENPH and RUN, however, continued to gain value in 2018. ENPH was the strongest performer across both years. 

### Refactoring the Script
The original version of the stock analysis script was functional, but left room for optimization. I refactored the code and considerably improved its execution time. Compare the run time of the original script vs. the refactored script.

{insert screen shot here}

The original script looped every row of the worksheet once for each stock in the dataset, outputting the results at the end of each loop. The refactored script instead outputs its results to arrays, and only needs to loop through the data once. This means it runs much faster.

## Summary
#### What are the advantages or disadvantages of refactoring code?
Refactoring code is time consuming work. Time that could be spent on a new project instead is used reviewing and editing existing code. But the effort to refactor may be worthwhile to improve the efficiency and logical structure of your code. Refactored code may be easier for another programmer to understand and it edit, and it may make a script run faster and use less memory.

#### How do these pros and cons apply to refactoring the original VBA script?

The original VBA script used in this project was functional, produced correct results, and ran in a short amount of time. But it only needed to output results for 12 stocks, and may not have performed as well with a larger data set. By refactoring the script to use fewer loops, its run time improved significantly. The resulting script is not only faster, but can likely handle a much larger dataset, making it more expandable.
