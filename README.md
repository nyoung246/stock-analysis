# Stock Analysis
# Overview of Project: 
The purpose of this analysis is to assist Steve in refactoring the original VBA script to increase its efficiency and user friendliness. The refactored script will also take into account all information from 2017 and 2018 worksheets to determine if these stocks are a good buy. It will generate the “Tickers”, “Total Daily Volume” and “Returns” shown on the “All Stocks Analysis” worksheet using the prompted year chosen.

## Results:

### Comparing Performance from 2017 and 2018
Comparing the stock performances between 2017 and 2018, as seen in the below image, the 2017 stocks all had positive returns with the top performers being DQ with the return of 199.4% and SEDG with the return of 184.5%. However, TERP underperformed with the negative return of -7.2%. In 2018, all stocks seem to have underperformed with negative returns with DQ having -62.6% return and HASI with -60.5% return. The two top performers for this year were RUN with 84.0% return and ENPH with 81.9%. Both ENPG and RUN both seem to be consistent with positive results for both 2017 and 2018 with “Total Daily Volumes” increasing by 173.92% and 87.82% respectfully.

![VBA_Challenge_2017](https://github.com/nyoung246/stock-analysis/blob/main/resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018](https://github.com/nyoung246/stock-analysis/blob/main/resources/VBA_Challenge_2018.PNG)

### Examples of the Code
Showing examples of the steps taken required to generate the 2017 and 2018 results, the below image is the first step from the refactored script. The “tickerIndex” was set to “0”. With this we are able to build upon the correct index for the following three output arrays “tickerVolumes”, “tickerStartingPrices” and “tickerEndingPrices”.

![Code_question_1](https://github.com/nyoung246/stock-analysis/blob/main/resources/Code_question%201.PNG)

For the next step shown in the below image, a FOR LOOP was created to initialize the “tickerVolumes” to “0” and another FOR LOOP was used to ensure that it would loop until the last row.

![ Code_question_2](https://github.com/nyoung246/stock-analysis/blob/main/resources/Code_question%202.PNG)

In next step shown below, we increased the volume for the current ticker by “tickerVolumes” with the variable and the addition of the ticker. Then an IF THEN statement was used to check if the current row is the first row with the selected “tickerIndex”. If the next row’s ticker did not match, then the “tickerindex” would increase, which was actioned using IF THEN statements.

![ Code_question_3](https://github.com/nyoung246/stock-analysis/blob/main/resources/Code_question%203.PNG)

As for the final step required shown in the below image, a for loop statement was used to display the “Ticker”, “Total Daily Volume” and “Return” results on the “All Stocks Analysis” worksheet.

![ Code_question_4](https://github.com/nyoung246/stock-analysis/blob/main/resources/Code_question%204.PNG)

### Comparing Execution Time
As for execution times of the original script and refactored script, I would say that there was an improvement in both the execution times. As seen in the below images, the original script from the 2017 data in the file Green_stocks, had an execution time of 1.011719 seconds and the refactored script from the file VBA_Challenge had an execution time of 0.1953125 seconds. The refactored script successfully ran 0.8164065 seconds faster.

![green_stocks_2017_run_time](https://github.com/nyoung246/stock-analysis/blob/main/resources/green_stocks_2017%20run%20time.PNG)

![VBA_Challenge_2017_run_time](https://github.com/nyoung246/stock-analysis/blob/main/resources/VBA_Challenge_2017%20run%20time.PNG)
As for the original script from the 2018 data in the file Green_stocks, had an execution time of 0.9804688 seconds and the refactored script from the file VBA_Challenge had an execution time of 0.1523438 seconds. The refactored script successfully ran 0.828125 seconds faster.

![green_stocks_2018_run_time](https://github.com/nyoung246/stock-analysis/blob/main/resources/green_stocks_2018%20run%20time.PNG)

![VBA_Challenge_2018_run_time](https://github.com/nyoung246/stock-analysis/blob/main/resources/VBA_Challenge_2018%20run%20time.PNG)

## Summary:

### Advantages of Refactoring Code
A few advantages of refactoring code is that you are already working with a base, so it is easy to find a starting point as opposed to writing the entire code again. Another advantage is to find any issues within the previous code written and to help execute the programs faster. This will then lead to time savings and more reader friendly codes for the next person to see. 

### Disadvantages of Refactoring Code
One disadvantage of refactoring codes is that because you are already working on an existing code, your thought process may be skewed towards to how the previous code was written. And another disadvantage is if the previous code was not organized with comments, it is more difficult and time consuming to find the issues within the code.

### How do these pros and cons apply to refactoring the original VBA script?
In the case for this assignment, one of the advantages of refactoring the original VBA script is to create a more efficient/faster code from what was already given. This way you can have a comparison on whether your refactored code was an improvement by looking at the run times. Another advantage of refactoring the original VBA script is that the comments were added at every step, so you can easily identify the steps taken to produce the results. One disadvantage is that you must write the entire script before attempting to run it. Therefore, you do not know if what you have already written would truly be more efficient until you have completed it.
