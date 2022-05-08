# stocks-analysis

## Overview of Project

### Background

Our client Steve originally requested an analytical report the green "DAQO" stock his parents invested in. Using VBA macros, we provided an analysis of the total daily volume and return percentage of the DAQO stock over a two-year period (2017 - 2018). Steve later expanded his request to provide an analysis of all the environmentally friendly green energy stocks, in order help his parents diversify his portfolio. The completed project provided Steve an macro button the returns an analysis of the desired stock upon input. Steve will now like to take this analysis and apply it to an even larger dataset.

### Purpose

With a larger dataset being introduced, the current must be refactored to improve efficiency and speed. Steve’s current macro was designed to run on the original dataset, which included a few dozen stocks. If we chose to run the same macro for a dataset that contains thousands of stocks, we may run into performance and speed issues. The purpose of this project is to refactor the current code and improve its performance on the current dataset in order to prepare the code for a larger dataset.

## Analysis

In order improve the efficiency of the code we first determined where changes can be made. One major influence for code speed is how much memory is being used. After, review of the code it was determined that the nested loops may cause the longest delay as loops require more memory during use.

Originally the program looped through the “tickers()” array.

![tickers_array](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/tickers_array.png?raw=true)
![tickers_array_loop](https://github.com/JoseEspinosaTello/stocks-analysis/blob/main/Recources/tickers_array_loop.png?raw=true)

