# Election Analysis 

## Overview
To help Steve do stock research for his parents, I have created a spreadsheet using Excel VBA code.  Within the spreadsheet, there is a button that will activate a user input box to separate the results by year.  In addition, due to the large volume of stock analysis data, I have added refactored code to take less steps, use less memory, and improve the logic to make it easier for future users to read.

## Results
As you can see from the images below the refactored code sped up the run time.

*Run time for 2017 ***before*** adding the refactored code:*

![2017_All_Stock_Analysis](https://user-images.githubusercontent.com/90974647/136658296-63c3e1c3-3b52-4087-a078-27fb4a1c5626.png)




*Run time for 2017 ***after*** adding the refactored code:*

![2017_ASA_Refactored](https://user-images.githubusercontent.com/90974647/136658393-3e78601b-e529-4e57-a38f-69bed0ca3034.png)




*Run time for 2018 ***before*** adding the refactored code:*

![2018_All_Stock_Analysis](https://user-images.githubusercontent.com/90974647/136658696-67950ca6-7935-4b2a-a0b8-ce3532e98a52.png)




*Run time for 2018 ***after*** adding the refactored code:*

![2018_ASA_Refactored](https://user-images.githubusercontent.com/90974647/136658699-282f559c-24cc-4533-ae55-16f9db530bdb.png)


**This shows a gain of .7 seconds!**

## Summary
By adding a timer, 
```
startTime = Timer
```
creating an index, and refactoring the original code set
```
For k = 0 To 11
Worksheets("All Stocks Analysis").Activate
Cells(4 + k, 1).Value = tickers(k)
Cells(4 + k, 2).Value = tickerVolumes(k)
Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
```
the improved script loops through the data one time to collect all of the information and improves performance by .7 seconds.
