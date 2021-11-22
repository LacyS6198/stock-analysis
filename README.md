# Module 2 Challenge - Stock-Analysis Using Excel VBA


## Overview of Project
The purpose of this project is to help our client analyze the performance of stocks in 2017 and 2018. The stocks were analyzed based on total daily volume and percentage returns over the year. 
- **Total Daily Volume** is equal to the sum of daily volume for a given ticker and year. 
- **Return** was calculated by dividing the ending stock price for a given ticker and year by its starting price; value is displayed as a percentage.

In addition to calculating the total daily volume and return of the various stocks, this workbook was created to be user-friendly and avoid manual calculations. VBA code was used to automate the calculations of each value. It also was developed to allow the user to reset the analysiz with the click of a button, as well as calculator either year's values at the click of a button. 

Finally, the original code was refactored to allow the analysis to run in a shorter period of time than the original code. This provides a better user experience by shortening the time of the analysis, as well as allows the workbook to handle even larger volumes of data. 

## Results

### Stock Performance
There was a significant difference in stock performance between 2017 and 2018. As shown in the below comparison, in 2017 the majority of stocks had a positive return for the year. In comparison, for 2018 the majority of stocks had a negative return. The average return in 2017 was 67.3% while the average return in 2018 was -8.5%. 

![2017_to_2018_StockPerformance](https://user-images.githubusercontent.com/93630042/142767828-08f1d8bb-6dd1-4545-a450-1dcf6b1e5d20.png)

#### Change in Daily Trading Volume
Between 2017 and 2018 there were changes in the daily trading volume totals. Seven (7) tickers had an increase in total daily trading volume while five (5) had a decrease over the same time period. 

While 7 tickers had an increase in daily volume, a total of 10 tickers had a decrease in return percentage. There was not a clear correlation between changes in daily trading volume totals and percentage return changes from 2017 to 2018. There were stocks that had a significant increase in total daily trading volume yet also a significant decrease in return. 
- Example: Ticker "DQ" had a 201% increase in trading volume (from 35,796,200 in 2017 to 107,873,900 in 2018) yet a 262% decrease in return (from 199.4% in 2017 to -62.6% in 2018)
- Example: Ticker TERP had a 9% increase in trading volum (from 139,402,800 in 2017 to 151,434,700 in 2018) yet only a 2.2% increase in return (from -7.2% in 2017 to -5.0% in 2018).

Based on the above two examples and an overall review of the 2017 and 2018 stock performances, we can see that the change in daily total volume did not seem to have a clear or direct impact on the percentage return.

#### Change in Return Percentage
Overall, 10 of the 12 stock tickers had lower percentage return in 2018 than in 2017; 2 had a higher percentage in 2018 than 2017. 11 of the 12 tickers had a positive return in 2017 while only 2 had a positive return in 2018. As a whole, the 2018 performance of stocks was worse than in 2017. 

#### Limitations
Further analysis would be required to determine why the stocks performed worse and why two performed better. The data used for analysis was limited to the ticker symbol, stock price throughout the day (open, close, low, high, adj close) and daily trading volume. It also only included the years 2017 and 2018. The following are suggestions for additional analyses that could be performed:
 - Historical stock data for previous 5 years plotted on a line chart to determine trend.
 - Economic factors between 2017 and 2018 that could have impacted stock performance. 
 - Overall market performance for 2017 and 2018 to determine if stocks were on-par with the market.
 - Category of stocks and average performance of stocks within each category for the given years to better determine if the performance was on-par for the stock type.

### Refactored Code Performance
After creating the original analysis workbook, the code was refactored to decrease calculation time and improve performance. Performance was measured based on run time of the calculations. In order to determine improvement in timing, the worksheet was cleared and the original code was ran for the 2017 analysis. The worksheet was then cleared again and the refactored code was ran for the 2017 analysis. The process was then repeated the year 2018. The difference in timing was significantly improved after the refactoring. 
- The 2017 analysis originally ran in 0.668 seconds. After refactoring, it ran in 0.102 seconds. This is an approximately 85% decrease in timing.
- The 2018 analysis originally ran in 0.766 seconds. After refactoring, it ran in 0.117 seconds. This is an approximately 85% decrease in timing. 

##### 2017 Run-Time Before Refactoring
![RunTime_2017_OriginalCode](https://user-images.githubusercontent.com/93630042/142784577-dc692260-ad32-421a-a7e6-2fcb66117eb0.png)

##### 2017 Run-Time After Refactoring
![RunTime_2017_RefactoredCode](https://user-images.githubusercontent.com/93630042/142784581-312f9955-eebe-400b-b68a-f3b5b6577ece.png)

##### 2018 Run-Time Before Refactoring
![RunTime_2018_OriginalCode](https://user-images.githubusercontent.com/93630042/142784584-78174600-eeea-4ae3-9626-b2451cd0acd1.png)

##### 2018 Run-Time After Refactoring
![RunTime_2018_RefactoredCode](https://user-images.githubusercontent.com/93630042/142784589-8edae95c-39e3-4e80-a65f-b6ee4e3692ae.png)


## Summary
