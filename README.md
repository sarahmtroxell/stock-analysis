# An Analysis of Stock Market Data

## Overview of Project
A fictional character, Steve, has requested help to analyze his parent's recent stock purchase. They purchased shares of DAQO New Energy Corporation (DQ); however, did no market research before making their decision. This project will analyze the performance of other stocks to determine if his parent's purchase was a wise financial decision based on each stock's annual return and daily volume. His parents want to include data from the entire stock market over the last few years, so the original VBA code built for Steve to analyze a small set of Green Energy stocks will need to be refactored to run faster.

### Purpose
The purpose of this challenge is to become comfortable coding in Excel VBA and practicing to solve a problem with multiple solutions. There are many different ways to solve a problem, but some are more optimized and run faster than others. By timing the execution of our original code with nested for loops against our refactored code, we can see that adding arrays and other improvements significantly reduced the run time of the refactored code. Refactoring other developer's code is also an important skill to practice because it makes you comfortable reading and working with code in a different format or approach that you may typically write. 

The following VBA methods were used to solve the challenge: 

1. Creating Macros
2. Arrays
3. Conditional Logic
4. Nested For Loops
5. Refactoring Code
6. User Form Design (Inputs, Message Boxes, Buttons)

## Results


# An Analysis of Stock Market Data

## Overview of Project
A fictional character, Steve, has requested help to analyze his parent's recent stock purchase. They purchased shares of DAQO New Energy Corporation (DQ); however, did no market research before making their decision. This project will analyze the performance of other stocks to determine if his parent's purchase was a wise financial decision based on each stock's annual return and daily volume. His parents want to include data from the entire stock market over the last few years, so the original VBA code built for Steve to analyze a small set of Green Energy stocks will need to be refactored to run faster.

### Purpose
The purpose of this challenge is to become comfortable coding in Excel VBA and practicing to solve a problem with multiple solutions. There are many different ways to solve a problem, but some are more optimized and run faster than others. By timing the execution of our original code with nested for loops against our refactored code, we can see that adding arrays and other improvements significantly reduced the run time of the refactored code. Refactoring other developer's code is also an important skill to practice because it makes you comfortable reading and working with code in a different format or approach that you may typically write. 

The following VBA methods were used to solve the challenge: 

1. Creating Macros
2. Arrays
3. Conditional Logic
4. Nested For Loops
5. Refactoring Code
6. User Form Design (Inputs, Message Boxes, Buttons)

## Results


### Original Code
The stock market data was given in an Excel sheet with a tab for 2017 data and 2018 data; the format is shown below. The To find each stock's annual return and total daily volume by year, I used the Ticker, Date, Open, Close, and Volume data. 
![Stock_Market_Data](../main/Resources/VBA_Challege_Data.png)

My original code created an array to hold the 12 ticker names. I then created a nested for loop, where counter i looped through each ticker stored in the array, and in the nested loop counter j looped through the rows of data. The nested loop is where I used conditional if statements and 3 different variables to set the amount of total volume, starting price, and ending price for each ticker that corresponded to counter i. The conditional statements are shown below:


```
     'loop through rows in the data
       For j = 2 To RowCount
        
        'Find the total volume for the current ticker
           If Cells(j, 1).Value = ticker Then
                   totalVolume = totalVolume + Cells(j, 8).Value
           End If
        
        'Find the starting price for the current ticker.
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                   startingPrice = Cells(j, 6).Value
           End If
        
        'Find the ending price for the current ticker.
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
           
       Next j
```

The last step of the macro was to output the data on the designated worksheet. My code previously formatted the data to have headers, so I activated the worksheet, referenced the cells I wanted each value to print to, and then assigned the value of each variable. To find the total return of each stock, I divided the ending price by the starting price and subtracted 1. This was done within the main for loop so I could capture the variable values only for each specific ticker before the code moved onto the next ticker and reassigned the variable values. 

Additionally, I assigned the Macros to buttons so users could analyze stock market data without needing to know how to use Excel Developer tools. A screenshot of the final analysis output, including conditional formatting to quickly identify high and low stock returns, is below: 

![VBA_Output](../main/Resources/VBA_Challenge_Output.png)


I ran the code two separate times to analyze stock market data from 2017 and 2018. Using the VBA timer function, I recorded the time my code took to complete the analysis. The recorded run times for each year's data are in the Message Box pop ups screenshotted below: 

![VBA_Original_2017](../main/Resources/VBA_Original_2017.png) ![VBA_Original_2018](../main/Resources/VBA_Original_2017.png)

### Refactored Code
Using the same data set as before, I was challenged to refactor the original code to execute the same results, but have a more simple format, run faster, and use less memory.

In the previous code, I used complicated nested for loops with two different counters. In my refactored code, I chose to store the variables for each ticker in an array. My code below shows the 3 arrays, where 12 corresponds to the amount of tickers. Throughout the single for loop, 12 will be replaced with a reference to a ticker index variable created. This eliminated the need of a nested for loop, which reduced the memory required.

```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

I also recognized that my original code printed the stock analysis for each ticker on the output sheet at the end of each loop. That process was timely, so my refactored code instead prints ad calculates the metrics at the end all at the same time. I believe this reduced the run time significantly. 
```
     'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
     For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
     Next i
```

I ran the final refactored code two separate times to analyze stock market data from 2017 and 2018. Using the VBA timer function again, I recorded the time my code took to complete the analysis so I could compare to my original code. The recorded run times for each year's data are in the Message Box pop ups screenshotted below: 

![VBA_Refactored_2017](../main/Resources/VBA_Challenge_2017.png) ![VBA_Refactored_2018](../main/Resources/VBA_Challenge_2018.png)

## Summary

### Comparison of Refactored Code
While both my original and refactored code for the VBA challenge completed the same task and gave the same output, the refactored code ran 0.488 seconds faster for 2017 data and 0.504 seconds faster for 2018 data. This may not seem like a lot of time, but when the code is applied to larger data sets, the gap can become significant. I also noticed that my refactored code was much easier to read back through and explain to others. The original code listed the tickers in an array, but printed the variables for each ticker within the for loop. The refactored code was able to optimize this by storing the values for each variable in an array that corresponded to each ticker, then printed the values all at once at the end of the code. Also since the original code used a nested for loop, there were two different counters for the code to store and run through. The refactored code was able to run using only one counter, which most likely reduced memory. My refactored code included the output headers and conditional formatting, where as the original code had a separate macro required to run to format the sheet. The only downside to refactoring my code was to learn how to use arrays properly in my code, since that was a new concept for me. In the future, I believe refactoring code will become easier as I become more comfortable with syntax, structure, and functions. 

### Key Takeaways
The clear advantage of refactoring code in general is to reduce the runtime and memory of your code. Depending on the size of data you are analyzing, the speed and memory required to run your code can become very important ad in extreme cases could crash the program. Additionally, refactored code tends to be easier for other developers to follow because its format is simpler and the functions are cleaner. If you wrote code for a job or client and it needed to be updated, you want others to be able to easily read your work and be able to collaborate without your explanation. The disadvantage of refactored code is that it may be time consuming for the developer. You can often reuse code formats from other projects to save time, but then may still have to refactor your code before finishing.
