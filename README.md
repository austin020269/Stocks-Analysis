# stock-analysis

UT Bootcamp Module 2 Challenge

## Project Overview
Use the Visual Basic Module within Excel to analyze and compare stock performance for datasets taken from the years 2017 and 2018.

## Resources
Data Sources provided to analyze and minipulate included:
- green_stocks.xlsm

Software utilized for this study included: 
- Visual Basic 
- Excel
- Personal GitHub account

## Analysis/Workflow/Results

### Deliverable 1: Refactor VBA Code and Measure Performance 

Specifically for this deliverable we did the following:
1. Download the challenge_starter_code.vbs file and rename it VBA_Challenge.vbs.
2. Create a folder called “Resources” to hold the run-time pop-up messages that you’ll screenshot after running refactored analyses for 2017 and 2018.
3. Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm.
4. Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
5. Use the steps below to add code where indicated by the numbered comments in the starter code file.
6. Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
7. Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
8. Create a for loop to initialize the tickerVolumes to zero.
9. Create a for loop that will loop over all the rows in the spreadsheet.
10. Inside the for loop, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.
11. Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices variable.
12. Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
13. Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
14. Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
15. Run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module (as shown in the images below). In your Resources folder, save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook.

Refactored portion of Code:

![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/refactored%20code.PNG)
![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/refactored%20code_2.PNG)

Outputs showing refactored code runs faster than the original for both years:

2017 original:

![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017_original.PNG)

2017 refactored:

![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017_refactored.PNG)

2018 original:

![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018_original.PNG)

2018 refactored:

![alt text](https://github.com/austin020269/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018_refactored.PNG)

### Deliverable 2: Written Analysis of Results  

Specifically for this deliverable we did the following:
1. Continue using the crypto_clustering.ipynb file from Deliverable 1 where you’ve already performed the preprocessing steps.
2. Using the information we’ve provided, apply PCA to reduce the dimensions to three principal components.
3. Create a new DataFrame named pcs_df that includes the following columns, PC 1, PC 2, and PC 3, and uses the index of the crypto_df DataFrame as the index (see below).


