# stocks-analysis

## Overview of Project

### Background

Our client Steve originally requested an analytical report the green "DAQO" stock his parents invested in. Using VBA macros, we provided an analysis of the total daily volume and return percentage of the DAQO stock over a two-year period (2017 - 2018). Steve later expanded his request to provide an analysis of all the environmentally friendly green energy stocks, in order help his parents diversify his portfolio. The completed project provided Steve an macro button the returns an analysis of the desired stock upon input. Steve will now like to take this analysis and apply it to an even larger dataset.

### Purpose

With a larger dataset being introduced, the current must be refactored to improve efficiency and speed. Steve’s current macro was designed to run on the original dataset, which included a few dozen stocks. If we chose to run the same macro for a dataset that contains thousands of stocks, we may run into performance and speed issues. The purpose of this project is to refactor the current code and improve its performance on the current dataset in order to prepare the code for a larger dataset.

## Analysis

One major influence for the speed of code is how much memory is being used while the code runs. After review of the code, it was determined that the nested loops may cause the longest delay as loops require more memory during use.

Originally the program looped through the “tickers()” array.

![tickers_array](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/tickers_array.png?raw=true)
![tickers_array_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/tickers_array_loop.png?raw=true)

Within main loop we created a nested loop that goes through each row and uses if-statements to find the total volume, starting price, and ending price of the current ticker value.

![nested_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/nested_loop.png?raw=true)

Once the nested loop ran through every row for the current ticker value, the program would return to the main loop and print those values to the cells for the current ticker value effectively ending the current loop and starting the loop over using the next ticker value.

![end_main_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/end_main_loop.png?raw=true)

As the program stores data and runs through the nest loop, it is also running and storing data from the main loop. Therefore, the nested loops are where an estimated high consumption of memory would occur. 

### Refactored code

The nested loop in our original code is where the bulk of the work is being done consuming the most energy. In order to eliminate the nested loop, we had to rework the way our arrays are handled. 
We began by initializing a tickerIndex as 0 and initializing the totalVolume, tickerStartingPrices, and tickerEndingPrices as arrays with an input of 11 equal to the size of the tickers() array.

![ticker_index_arrays](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/ticker_index_arrays.png?raw=true)

Open a loop that sets the value of all 12 items in the list to 0 and then close the loop.

![total_volume_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/total_volume_loop.png?raw=true)

Next, we will loop through all rows and apply the proper values to our new arrays. Instead of determining these values and setting them during the nested loop like the old code, the refactored code will store the values in the arrays until the value is call in the next loop.

![refactored_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/refactored_loop.png?raw=true)

The final loop will loop throw the now fully populated arrays and call the value of the current item in the list and assign the value to the desired cells in the excel form.

![ref_final_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/ref_final_loop.png?raw=true)

Overall, removing the nested loops and turning the totalVolume, tickerStartingPrices, and tickerEndingPrices into arrays, will allow us to store the values we determine within the arrays outside the loops. This gives us the freedom to create three loops that run at separate times and require less memory consumption, therefore increasing the speed of the program.

## Results

The initial test of the code showed incredible results. The first test of the original code returned 0.5507813 seconds for 2017 and 0.5507813 seconds for 2018.

![2017_reg_1](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/2017_reg_1.png?raw=true) ![2018_reg_1](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/2018_reg_1.png?raw=true)