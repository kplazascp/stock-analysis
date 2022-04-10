# An Analysis of Green Stocks
In this written analysis for Steve, we are helping Steve and his parents to analyze different stock options with VBA, for Steve parents that are passionate about alternative energy production.

## Overview of Project

Steve´s parents decided to invest their money into DAQO New Energy Corp, Steve however wants to diversify their portfolio by searching for different green stock options. For that reason, you will find the following analysis focusing on green stocks for the years 2017 and 2018, where we help Steve Analyze the different green stock options available.

### Purpose

The purpose of this analysis is to help Steve and his parents to diversify their investment. Steve´s parents have already invested all their money in DQ stocks, our role in this analysis is to help Steve get a better understanding of how the performance of DQ has been in 2017 and 2018, and also how is the Volume and Return from other 11 green stock options available, so that him and his parents can invest wisely.
This analysis was made by creating a Visual Basic for Applications (VBA) script that will allow us to run all the green stocks data from 2017 and 2018, so that we can analyze a lot of data very fast and automate the metrics that we will follow for all the tickers, instead of performing the analysis one ticker at a time.
This will also help Steve use it with any other stock that he wants in the future.

## Results
### A more efficient code
To be able to help Steve include the entire stock market over the last few years , we refactored the code (made the code more efficient), so that it runs faster (it takes less time to execute) and will not be affected in the future when Steve wants to add more data. 

With the initial code we had, before it was refactored, our code run in 0.9765625 seconds for 2017 and 1.007813 seconds for the year 2018.

![Original_code_run_2017](https://github.com/kplazascp/stock-analysis/blob/main/Time_to_run_2017.PNG)
![Original_code_run_2018](https://github.com/kplazascp/stock-analysis/blob/main/Time_to_run_2018.PNG)

To be able to make the code more efficient, we created a tickerIndex variable set to zero before iterating all the rows
`tickerIndex = 0`

Then we created three output arrays: TicketVolumes, tickerStartingPrices and tickerEndingPrices

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

After this, it is time to create a `for` loop that will initialize the `tickerVolumes` to zero

```
For i = 0 To 11
        tickerVolumes(i) = 0
Next i
```
    
And created a `for` loop that will loop over all the rows in the spreadsheet, in this `for` loop, we will write a script that increases the current `tickerVolumes` (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

```
''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

We then wrote an `if-then` statement to check if the current row is the last row and if it is, assign the current closing price. 

```
'3b) Check if the current row is the first row with the selected tickerIndex. 
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
  tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
```

Then, we wrote a script that increases the `tickerIndex` if the next row´s ticker doesn´t match the previous row´s ticker

```
'If the next row’s ticker doesn’t match, increase the tickerIndex.
'3d Increase the tickerIndex. 
  If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      tickerIndex = tickerIndex + 1
  End If
Next i
```

Finally, we used a for loop to loop through all the arrays to output `Ticker`, `Total Daily Volume` and `Return`

```
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return. 
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i) 
        Cells(4 + i, 2).Value = tickerVolumes(i) 
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
   Next i
```

The result of the refactored code was that the year 2017 now runs in 0.15625 second and the year 2018 runs in 0.1484375 seconds.

![VBA_Challenge_2017](https://github.com/kplazascp/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018](https://github.com/kplazascp/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

### Stocks Results

![Results_2017](https://github.com/kplazascp/stock-analysis/blob/main/Results_2017.PNG)

With the refactored code we can evidence that for the year 2017, 11 out of 12 tickers had a positive return, except TERP that had -7.2%. The highest return (in percentage) was achieved by DQ, but the highest Total Daily Volume was achieved by SPWR with 782,187,000 and 23% of Return, and the second highest volume was FSLR with 684,181,400 and 100% of Return.

![Results_2018](https://github.com/kplazascp/stock-analysis/blob/main/Results_2018.PNG)

For the year 2018, we can evidence that it was not a very good year for green stocks, were 10 out of 12 stocks had a negative return. The highest return was of RUN with 84% and 502,757,100 Total daily Volume and ENPH with 82% and 607,473,500. 
For DQ, that was the particular interest of Steve´s family the return for 2018 was of -62.6% with a Total Daily Volume of 107,873,900.

## Summary
-The advantages of refactoring code are that the more efficient code that you have, the better performance it will have. If we have the need to analyze thousands of data, we will not be able to do so with an inefficient code.

-The disadvantages of refactoring code it that it takes more time, and time is an essential factor. When you have a code that runs perfectly, sometimes there are people that don’t see the need on refactoring their code. However, as we mentioned in the previous point, an inefficient code will only buy you some time and then it will stop working as you intended it to do.

-To pros of refactoring the original VBA script is that we can see hand on hand the effect in time and resources that a refactored code has compared to the original code. Also, it gives us the opportunity to practice and apply all that we have learned in order to create a more optimized code.

-The cons that apply to refactoring the original VBA script are that sometimes you have to forget about certain things that worked in your previous code (some, not all) that will not work in your new code. So, you will have some parts that you can adapt to your new code, but you have to see this new code as it is, a new code, not the original code.

If you are having trouble refactoring your code, I attach here the link for the excel spreadsheet, where you can find the VBA script [VBA_Challenge](https://github.com/kplazascp/stock-analysis/blob/main/VBA_Challenge.xmls.xlsm).

