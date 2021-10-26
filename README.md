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

Next, a nested for loop was written to loop through the stock performance data, calculate the total daily volume, starting price, ending price, and return for one ticker, or alternative energy company, then output the results for that one ticker before looping through the data all over again for each consecutive ticker individually.

[Original_For_Loops.png]

#### Execution Times

The execution time for the original script was approximately 0.309 seconds on 2017 data, and 0.297 seconds on 2018 data.

[Original_Timer_2017.png]           [Original_Timer_2018.png]

#### Final Output

The original script resulted in the following final outputs for 2017 and 2018 consecutively:

### The Refactored Script

While the original script was able to calculate and output the desired stock performance data, it was refactored in order to reduce the execution time to run the code. Refactoring the script would be most beneficial in the scenario that hundredes of thousands of lines of data, or more than 11 companies, would need to be looped through. Similar to the original script, an array of all tickers was initialized. But instead of using a nested for loop to calculate and output stock performance for each company individually, new arrays were initialized to store the calculations for the daily volume, starting price, ending price, and return of each company in its own variable.

#### New Variables and Arrays
In the refactored script, a new variable was introduced called "tickerIndex." Instead of using "i" for the first for loop, tickerIndex would be used to indicated which ticker, or company, stock performance data was being calculated for. Three new arrays were also introduced, consisting of "tickerVolumes(12)", "tickerStartingPrices(12)", and "tickerEndingPrices(12)." These arrays would hold the calculated daily volume, starting price, ending price, and return for all 11 companies.

[Refactored_Variables.png]

#### Refactored Nested For Loops
The nested for loops from the original script were refactored to perform the same calculations, but instead of outputting the data after each loop, the calculations were stored in the new arrays, ready to be used anywhere outside of the nested for loops.

[Refactored_For_Loops.png]

#### Output For Loop
An additional for loop was added outside of the nested for loops, which would output the results for daily volume, starting price, ending price, and return for all 11 companies by calling the variables stored inside of the new arrays.

[Refactored_Output_For_Loop.png]

#### Formatting
Formatting for the All Stocks Analysis sheet was also added to the refactored script, which involved the following formats:
1. Bold the text of the header row
2. Add a bottom border to the header row
3. Separates digits with commas and displays a trailing zero for the Total Daily Volume column
4. Makes a single-digit percentage for the Return column
5. Auto-fits the data in the Total Daily Volume column
6. Conditionally formats the Return column by looping through the column and formatting the cell based on the sign of the value (green fill for positive or red fill for negative)

[Refactored_Formatting.png]

#### Execution Times

The execution time for the refactored script was approximately 0.102 seconds on 2017 data, and 0.094 seconds on 2018 data, significantly faster than the original script.

[VBA_Challenge_2017.png]      [VBA_Challenge_2018.png]

## Summary
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
