# Stock Analysis

## Overview of Project:
The purpose of this project was to find the most profitable green energy stocks and compare them to "DQ" from the years 2017 and 2018 to help Steve create a more diverse portfolio for his parents and educate them to make a more financially sound decision on not one just based on nostalgia.

## Analysis Results:
###### Stock Performance
According to the analysis "DQ" had the greatest percent return in 2017 compared to the rest of the stocks in the same year which shows good growth, although it did not sell the most in volume. The stock"SPWR" performed the best in 2017 with total volume of $782,187,000 with an increase in return of 23%. The only stock to show a loss in return was "TERP" with -7.2%. In 2018 "DQ" dropped the most with -62.6% loss in return. The only two stocks to show an increase in return was both "ENPH" and "RUN". Taking the previous year into account both "ENPH" and "RUN" performed better then "DQ" overall.

###### Execution Times
The original script took no less than one second to return the results for 2017. The refactored code performed much faster taking [about .7 seconds](https://github.com/67owilliams/Stock-Analysis/blob/main/VBA_Challenge_2017.png) to return the same results as well as formatting the cells. The same can be seen when comparing the times between the original and refactored code for 2018. The original takes about one second while the refactored code takes [about .7 seconds](https://github.com/67owilliams/Stock-Analysis/blob/main/VBA_Challenge_2018.png) to run.

## Summary
###### Pros & Cons: Refactoring Code
Some pros of refactoring code is to reduce the time it takes for the script to run. Another plus to refactoring code is making it easier to read and understand. On the downside refactoring code may introduce new bugs that were not there in original code.

###### Original vs. Refactored VBA Script
In the original code the variables startTime and endTimes were initiated as the Double type. 
The same for startingPrice and endingPrice that were set as Double creating larger values to be returned as results.

  Dim StartTime As Double
  Dim EndTime As Double

  Dim StartingPrice As Double
  Dim EndingPrice As Double

To cut down a bit on the time it takes to run the refactored code both startTime and endTime are initialised as the Single type. Also the data to be returned in the "All Stocks Analysis" worksheet are stored in arrays initialised as the Long, and single types.


  Dim StartTime As Single
  Dim EndTime As Single

  Dim tickerVolumes(12) As Long
  Dim tickerStartingPrices(12) As Single
  Dim tickerEndingPrices(12) As Single
