# Stock Analysis

## Overview of Project

The purpose of this analysis was to observe the [stock performance data](https://github.com/fobordo/stock-analysis/blob/main/VBA_Challenge.xlsm) between 2017 and 2018 of various alternative energy companies. In this analysis, the total daily volumes and returns for 11 alternative energy companies, such as DAQO New Energy Corporation, were calculated in order to determine which companies performed best and would be most lucrative to invest in. We used Excel and VBA to perform these calculations.

## Results
Two different VBA scripts, the original and the refactored, were used to compare the stock performance betewen 2017 and 2018 of the 11 alternative energy companies. Both scripts prompt input from the user to indicate which year/worksheet, 2017 or 2018, they want to calculate the stock performance for. While both scripts resulted in the same stock performance outputs, the refactored script was able to run much faster than the original.

### The Original Script
Originally, a longer VBA script was used to calculate and output the total daily volumes and returns of the 11 alternative energy companies. 

#### Array of All Tickers
First, an array of all tickers was initialized, as seen in the screenshot below. 

![Original Tickers Array](/Resources/Original_Tickers_Array.png)

The tickers array was given an index of 12 to store the ticker symbols of all 11 alternative energy companies. 

#### Nested For Loops

Next, a nested for loop was written to loop through the stock performance data in the sheet of the specified year (2017 or 2018). After one loop, the script would calculate the total daily volume, starting price, ending price, and return for one ticker, then output the results for that ticker onto the All Stocks Analysis sheet before looping through the data all over again for each of the 11 tickers until the end of the for loop.

![Original For Loops](/Resources/Original_For_Loops.png)

#### Formatting
The original script also included formatting for the All Stocks Analysis sheet (which stayed the same for the refactored script), including the following formats:
1. Bold the text of the header row
2. Add a bottom border to the header row
3. Separates digits with commas and displays a trailing zero for the Total Daily Volume column
4. Makes a single-digit percentage for the Return column
5. Auto-fits the data in the Total Daily Volume column
6. Conditional formatting for the Return column. A for loop loops through the column and formats the cell based on the sign of the value (green fill for positive or red fill for negative)

![Original Formatting](/Resources/Original_Formatting.png)

#### Final Output

The original script resulted in the following final outputs for 2017 and 2018 consecutively:

![Original Outputs 2017](/Resources/Original_Outputs_2017.png)           
![Original Outputs 2018](/Resources/Original_Outputs_2018.png)

#### Execution Times

The execution time for the original script was approximately 0.320 seconds on 2017 data, and 0.336 seconds on 2018 data.

![Original Timer 2017](/Resources/Original_Timer_2017.png)           
![Original Timer 2018](/Resources/Original_Timer_2018.png)  

### The Refactored Script

While the original script was able to calculate and output the desired stock performance data, it was refactored in order to reduce the execution time to run the code. Refactoring the script would be most beneficial in the scenario that hundredes of thousands of lines of data, or more than 11 tickers, would need to be looped through. Similar to the original script, an array of all tickers was initialized. But instead of using a nested for loop to calculate and output the stock performance for each ticker after each loop, new arrays were initialized to store the calculations for the daily volume, starting price, ending price, and return of each ticker in its own variable.

#### New Variables and Arrays
In the refactored script, a new variable was introduced called "tickerIndex." Instead of using "i" for the first for loop, tickerIndex would be used to indicate which ticker stock performance data was being calculated for. Three new arrays were also introduced, consisting of "tickerVolumes(12)", "tickerStartingPrices(12)", and "tickerEndingPrices(12)." These arrays would hold the calculated daily volume, starting price, ending price, and return for all 11 tickers.

![Refactored Variables](/Resources/Refactored_Variables.png)

#### Refactored Nested For Loops
The nested for loops from the original script were refactored to perform the same calculations, but instead of outputting the data after each loop, the calculations were stored in the new arrays, ready to be used anywhere outside of the nested for loops.

![Refactored For Loops](/Resources/Refactored_For_Loops.png)

#### Output For Loop
An additional for loop was added outside of the nested for loops, which would output the results for daily volume, starting price, ending price, and return for all 11 tickers by calling the variables stored inside of the new arrays.

![Refactored Output For Loop](/Resources/Refactored_Output_For_Loop.png)

#### Formatting and Final Output
The formatting code stayed the same in the refactored script, which resulted in the following final outputs for 2017 and 2018 consecutively:

![Refactored Outputs 2017](/Resources/Refactored_Outputs_2017.png)
![Refactored Outputs 2018](/Resources/Refactored_Outputs_2018.png)

As seen in the screenshots above, the final outputs for the original script and refactored script were exactly the same. The only different output between the two scripts was the execution times to run the code.

#### Execution Times

The execution time for the refactored script was approximately 0.102 seconds on 2017 data, and 0.094 seconds on 2018 data, approximately 0.2 seconds faster than the original script.

![VBA Challenge 2017](/Resources/VBA_Challenge_2017.png)
![VBA Challenge 2018](/Resources/VBA_Challenge_2018.png) 

## Summary
### The Advantages and Disadvantages of Refactoring Code
The advantage of refactoring code is that it improves the design of the code structure, making it easier to maintain and manipulate. Any developer can easily go through and make modifications as the data they are analyzing changes or grows. Other advantages include making the code easier for future users to read, understand, and identify bugs. Refactoring code also decreases the time it takes to run the code because it takes fewer steps to perform the same functions, and uses less memory.

While there are many advantages to refactoring code, some disadvantages exist too. A few disadvantages include running out of time in refactoring the code if there are time constraints on finishing a project, or inadvertantly introducing new bugs that didn't exist in the original code. Further, if the original code is already well written and runs efficiently, sometimes it may be non value-added to refactor the code if there are no plans to revisit or grow the project in the future.

### The Pros and Cons of Refactoring the Original VBA Script
The pros of refactoring the original VBA script are that if we obtained more data on the stock performance of more tickers from different years, we could easily change the arrays and initialized tickers to accomodate the new data. We would not have to redesign the code structure, eliminating the chances of encountering new bugs. It would also be beneficial to have faster refactored code in the case that we must analyze hundreds of thousands of rows of data, decreasing the execution times to run the code and minimizing memory usuage.

The cons of refactoring the original VBA script are that if this was a one-time project that would not be revisited in the future, which in this case it is, and it were for a big company, refactoring the original code that was working just fine would have been a waste of time and resources with little return. The difference in execution times between the original and refactored VBA script, approximately 0.2 seconds, was hardly noticeable since we were only analyzing 3000 rows and 11 tickers. While the faster refactored code would be beneficial for larger datasets, in this case, it was not significantly value-added to our project.
