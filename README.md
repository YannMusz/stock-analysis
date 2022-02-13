# stock-analysis

## Overview and Purpose
Due to the diminishing returns of the use of fossil fuels and the movement towards alternative methods of producing energy such as hydroelectric, wind, geothermal, bioenergy and solar energy, we are asked to create a project using Visual Basic for Applications (VBA) to analyze hundreds of green stocks during the years 2017 and 2018.

## Purpose
The main purpose of this project was to compare stock performance for 2017 and 2018 by using two scripting methods (Original and refactored) and determine which scripting method is the most efficient at providing us with the end result.

## Results

By running our macros, we were able to quickly and efficiently figure out what stocks provided a positive return and which stocks provided a negative one.

![All Stocks 2017](/Resources/All_Stocks_2017.PNG)  ![All Stocks 2018](/Resources/All_Stocks_2018.PNG)
 
 The Two tables above provide a direct comparison of the performance of greent stocks for both 2017 and 2018. The Ticker column shows the stock symbol, the total daily volume shows how a particular stock was traded in one day and the return column indicates the percentage change over that period of time. The return column is color coded to better indicate which stock was profitable vs not profitable. We can clearly see that 2017 was a good year for green stocks as all but one stock (TICKER : TERP) had positive returns. It is a stark contrast compared to 2018 where they were not as profitable as only two stocks (Tickers: ENHP and RUN) had negative returns.
 
 ## Summary
 We ran the analysis using our VBA scripts that was later altered into a refactored script for better performance and efficiency. Although the original code which contained nested for loops worked, the refactored code was implemented to more efficiently loop through all data provided at once. The runs times were calculated and the refactored code showed fast improvements in term of scripting time as shown below
 
 ![VBA Challenge 2017](/Resources/VBA_Challenge_2017.png) ![VBA Challenge 2018](/Resources/VBA_Challenge_2018.png)
 
 ### My Conclusions on the project
 The advantage that refactoring provided over the original code was quite noticeable in terms of efficiency. What was not efficient was having to go back and re-edit most of the originaly code as you may run into mutiple errors trying to debug. Although tedious, in the end refactoring turned out to be more organized and structured in terms of data presentation.
 
