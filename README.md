# Vanderbilt University Data Analytics Bootcamp Module 1: Refactoring _VBA_
## Project Overview
The Vanderbilt University Data Analytics Bootcamp is a streamline of coursework that exposes students to a broad set of technical applications and data management concepts, strengthening their problem solving skills in data-driven contexts. Module 2 of the course focuses on streamlining analyses in Microsoft Excel with subroutines in Visual Basic, focusing on creating readable code, static and conditional fomatting, for loops and nested for loops, buttons, and measuring code performance.  

This module culminated in a challenge tasking students to refactor a subroutine highlighting stock performance (Sub AllStocksAnalysis) so that the subroutine would run more quickly.  Processing speed is tracked with a timer, which starts after the subroutine is initiated by clicking the "Run Analysis" button, when the user inputs the time period framing the analysis (year 2017 or 2018), and ends when the subroutine is finished, after which the run time is confirmed via message box.

The challenge will exemplify how code can be better streamlined with refactoring, and yield opportunity to analyze why the refactoring was effective, informing strategy for building subroutines. In addition, students are challenged to write and analyzeclear code.

## Results
The subroutine was refactored by adding arrays that could house the data for all stocks (Volume, Starting Price, Ending Price).  This made it possible to restructure for loops to perform fewer operations in every iteration.  

For example, the original subroutine _both_ 1) found values to assign to a single ticker (Total Volume, Starting Price, and Ending Price) and 2) printthe data in the "All Stocks Analysis" for the single stock in the worksheet with each iteration of i.

<img src="Resources/GetTickerData_AndPrint.png" alt="your alt text" width="500"/>
[Resources/GetTickerData_AndPrint..png](LINK)

Once refactored, finding ticker data (steps 2a - 3c) and printing the data (steps 4) could be separated into different for loops.  All data gathering (assign starting and ending prices, totaling trading volume) for all stocks could be processed in the for loop in sections 2a - 3c , and printing of the data for all stocks could be processed in separate loop, in section 4.

<img src="Resources/GetTickerData_Print_Separate.png" alt="your alt text" width="500"/>
[Resources/GetTickerData_AndPrint..png](LINK)

Before refactoring, the run time for the subroutine was .35 seconds for 2017 and .34 seconds for 2018.

<img src="Resources/AllStocksAnalysis_2017Results.png" alt="your alt text" width="1000"/>
[Resources/AllStocksAnalysis_2017Results.png](LINK)
<img src="Resources/AllStocksAnalysis_2018Results.png" alt="your alt text" width="1000"/>
[Resources/AllStocksAnalysis_2018Results.png](LINK)

After refactoring, the run time for the subroutine was .28 seconds for 2017 and .28 seconds for 2018.  
<img src="Resources/VBA_Challenge_2017.png" alt="your alt text" width="1000"/>
[Resources/VBA_Challenge_2017.png](LINK)
<img src="Resources/VBA_CHallenge_2018.png" alt="your alt text" width="1000"/>
[Resources/VBA_CHallenge_2018.png](LINK)

Both analyses ran roughly 20% more quickly.   It is likely that breaking the for loops into more efficient iterations yielded the improved processing speeds.

##