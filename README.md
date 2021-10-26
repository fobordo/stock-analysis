# Stock Analysis

## Overview of Project

The purpose of this analysis was to observe the [stock performance data](https://github.com/fobordo/stock-analysis/blob/main/VBA_Challenge.xlsm) between 2017 and 2018 of various alternative energy companies. In this analysis, the total daily volumes and returns for 11 alternative energy companies, such as DAQO New Energy Corporation, were calculated in order to determine which companies performed best and would be most lucrative to invest in. We used Excel and VBA to perform the calculations in this analysis.

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Two separate VBA scripts, the original and the refactored, were used to compare the stock performance betewen 2017 and 2018 of the 11 alternative companies. Both scripts prompt input from the user to indicate which year, 2017 or 2018, they want to calculate the stock performance for. While both scripts resulted in the same stock performance outputs, the refactored script was able to run much faster than the original.

### The Original Script
Originally, a longer VBA script was used to calculate and output the total daily volumes and returns of the 11 alternative energy companies. 

#### Array of All Tickers
First, an array of all tickers was initialized, as seen in the screenshot below. 

[Original_Tickers_Array.png]

The tickers array had an index of 12 to represent the 11 alternative energy companies. 

#### Nested For Loops

Then, a nested for loop was written to loop through the stock performance data, calculate total daily volume and return, and output the results for each ticker, or alternative energy company, individually. 

[Original_For_Loops.png]

#### Execution Times

The execution time for the original script on 2017 data was approximately 0.309 seconds, and approximately 0.297 seconds on 2018 data.

[Original_Timer_2017.png]           [Original_Timer_2018.png]

### The Refactored Script
#### Array of All Tickers
#### Nested For Loops
#### Execution Times

## Summary
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
