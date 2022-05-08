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
