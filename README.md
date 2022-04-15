# Stock Analysis

## Overview of Project
We were previously given an [Excel workbook](./resources/VBA_Challenge.xlsm) containing two years of stock data for twelve specific stocks. In the worbook, we created several Visual Basic for Applications (VBA) macros to help display the data, but the code runs somewhat slowly. We want to know if the same results can be achieved in less time. So we will refactor the code in an attempt to do so.

## 

### Original Attempt
The original VBA code (found in Module1 of the Visual Basic Editor) involved making a pass through every stock transaction[^1] in the data to compile data for a given stock ticker code ("ticker"), and then repeating the process for each ticker in the set. The time required to complete the task for years 2017 and 2018 was around .559 and .590 seconds, respectively, as seen in the timer dialog screenshots below:

![2017 original attempt](./resources/green_stocks_2017_timer.png) ![2018 original attempt](./resources/green_stocks_2018_timer.png)

[^1]: In order for the compiled statistics to be accurate, the stock transactions must first be sorted by ticker, and then by date.

### Refactored Attempt
The code was rewritten so that the stock-ticker statistics could be held in arrays in RAM, which would then be transferred to the worksheet after all statistics had been compiled


the statistics can be compiled in a single pass[^1] 

## Results
[Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.]

## Summary
[In a summary statement, address the following questions:]
1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?
