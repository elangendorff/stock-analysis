# Stock Analysis

## Overview of Project
We were given an [Excel workbook](./resources/VBA_Challenge.xlsm) containing two years of stock data for twelve specific stocks. In the workbook were several Visual Basic for Applications (VBA) macros to help display statistics about the stock data, but the code ran somewhat slowly. We want to know if we can obtain the same results, but faster. So we will refactor the code in an attempt to do so.

## The Code

### Original Attempt
The original VBA code (found in Module1 of the Visual Basic Editor) involved making a pass through every stock transaction in the data[^1] in order to compile statistics for a given stock ticker code ("ticker"). The data was then output to an available line on the output worksheet, and the process was repeated for each ticker in the set.

[^1]: Note that, in order for the compiled statistics to be accurate, the stock transactions first had to be sorted by ticker, and then by date (ascending).

![Original attempt main-action loop](./resources/green_stocks_main-action_loop.png)

### Refactored Attempt
The code was rewritten (in Module2 of the Visual Basic Editor) so that the stock-ticker statistics could be held in arrays in RAM.

![ticker statistics arrays](./resources/VBA_Challenge_array_creation.png)

In addition, the main action of the statistics' compilation was altered so that all stocks' statistics could be compiled in a single pass[^1]. Rather than scan the entire data set for each ticker—as the original attempt did—the refactored version instead stores each statistic's data in an array cell whose index corresponds to the stock in question. When the system begins to encounter transactions for a new ticker, it simply changes the index it uses.

![Refactored attempt main-action loop](./resources/VBA_Challenge_main-action_loop.png)

After the statistical compilation pass completes and statistics have been compiled in the statistical arrays for all stock tickers, the contents of the arrays are then transferred to the worksheet.

![ticker statistics output loop](./resources/VBA_Challenge_output_loop.png)

## Results
Both versions of the code produced the same output[^2] (as desired), as seen in the following screenshots:

[^2]: The original attempt formatted the output slightly differently (found [here](./resources/green_stocks_2017.png) and [here](./resources/green_stocks_2018.png)). Importantly, though, the actual data—the contents of the table—is the same between both versions.

| **2017** | **2018** |
| --- | --- |
| ![2017 stock output](./resources/VBA_Challenge_2017.png) | ![2017 stock output](./resources/VBA_Challenge_2018.png) |

The refactored version of the code, however, produced its results in a much shorter time than the original version did. The original code required .56 and .59 seconds, respectively, to produce the 2017 and 2018 data, whereas the refactored code required only .19 seconds, each.

| | **2017** | **2018** |
| --- | --- | --- |
| **Original** | ![2017 original attempt](./resources/green_stocks_2017_timer.png) | ![2018 original attempt](./resources/green_stocks_2018_timer.png) |
| **Refactored** | ![2017 refactored](./resources/VBA_Challenge_2017_timer.png) | ![2018 refactored](./resources/VBA_Challenge_2018_timer.png) |

## Summary
Refactoring code can be a difficult process. It takes time and effort. If done well, however, it can produce code that runs better and faster, and is easier to read and to maintain.

We see these outcomes in this very exercise, most notably in the main-action loops: the refactored, single-pass loop (with `if-then` conditional checks inside) is much easier to understand than the nested `for` loop that appears in the original, and the refactored code also finishes in about a third the time of the original version.

We could even simplify the process, further: since the last two conditionals in the refactored main-action loop have the same condition ("Is the current row the last one for this ticker?"), their instructions could also be combined into the same `if-then` body, which would simplify and speed up the code even further. (Although with a change that simple, it will likely have very little noticeable effect unless the data set used becomes _extremely_ large.)

Whether or not making this change is worth the time and effort is, in the end, a judgment call that will be based on each user's, programmer's, or organization's particular needs and available resources.
