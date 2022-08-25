# stock-analysis
Module 2 VBA
## Overview of Project
### Background

With our help, Steve can analyze 12 green energy stocks, including DQ, to help his parents and new clients diversify their portfolios. Steve uses VBA and macros to perform complex logic and calculations to reduce errors, increase efficiency and speed at the click of a button. He now wants to expand his analysis to include the entire stock market over the last few years.

### Purpose

We’ll refactor Steve’s workbook to loop through all the data one time in order to collect the same information previously to increase the VBA script run time.

## Results

### Performance
The green stocks performed substantially better in 2017 compared to 2018. Formatting the positive returns in green and negative in red really helps illustrate that. 

![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/Screenshot%20All%20Stocks%20(2017).png)
![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/Screenshot%20All%20Stocks%20(2018).png)

### VBA Script Run Time
The execution times from the refactored script was 0.54 seconds faster than the original script. It would be really interesting to calculate how much time is saved returning thousands of stocks over 5 or 10 years.

![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/VBA_Challenge_2017%20AllStocksAnalysis.png)
![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/VBA_Challenge_2018%20AllStocksAnalysis.png)
![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

### VBA Original versus Refactored
The following screenshots show the difference between original Sub AllStocksAnalysis versus the Refactored Sub that enabled the run time to be reduced.

**Original**

![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/Sub%20AllStocksAnalysis().png)

**Refactored**

![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/Sub%20AllStocksAnalysisRefactored()1.png)
![TBrickey](https://github.com/TBrickey/stock-analysis/blob/main/Resources/Sub%20AllStocksAnalysisRefactored()2.png)

## Summary
### Advantage & Disadvantage
The advantage of refactoring is recycling old code to create the product faster. I can see the benefit of not having to “reinvent the wheel” every time to create a similar product. The disadvantage I encountered using the original module code, was I became careless and sloppy copying over the original script. Because I had so many problems, I started wondering if it would have been fastest to start from scratch.
