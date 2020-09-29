# Green Stock-analysis

## Overview of Project: 

The purpose of this analysis is to offer a clear and simple way for our friend Steve to quickly analyze stock data for different years and to do it in as efficient a manner as possible. In this case when we refactored the code we removed the nested for loop.

## Results: 

Our analysis has shown that with the exception of TERP stock all of these stocks did not do as well in 2018 as they did in 2017. The stocks ENPH & RUN both still had positive growth in 2018 so those would be the best stock to invest in for the long term. Also the stock DQ that Steve's parents had previously invested in had the highest returns in 2017 but the most losses on 2018 so they are very volatile and not good for long term investments. 

The refactored code was significantly faster to run. 

Unrefactored results for 2017 returned in `1.3359938` seconds while the refactored code took `0.3125` seconds. For the year 2018 the unrefactored code took `1.429688` seconds and the refactored code returned results in `.021875` seconds. Looping through the main data one but introducing a separate For loop to move through the tickers made a significant improvement in the time it took to run the code. 

```
    For i = 0 To 11
        
        tickerVolume(i) = 0

      Next i
```

This was made possible by also adding a variable called tickerIndex:

```
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
```
That we could increment everytime a change of ticker was detected.
```
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerIndex = tickerIndex + 1
            
            End If
```

### 2017 Results with Refactored Code:

![Refactored Results & Run Time for 2017](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

### 2018 Results with Refactored Code:

![Refactored Results & Run Time for 2018](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

### 2017 Results with Unrefactored Code:

![Un-Refactored Results & Run Time for 2017](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2017_unrefactored.png)

### 2017 Results with Unrefactored Code:

![Un-Refactored Results & Run Time for 2018](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2018_unrefactored.png)


## Summary: 

### What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code can include greatly increased efficiency of the code, having fewer steps or improving the logic within the code.

The disadvantages include the possibility that you will have invested a large amount of time and not actually improved the efficiency or anything else about the code. You are also spending time on a task that is already being completed successfully so you will need to weigh the time it will take against other tasks that need to be completed.

### How do these pros and cons apply to refactoring the original VBA script?

The disadvantage of refactoring this code is that it is more complicated which might make it harder to edit or improve later. However, the results of refactoring the code during this challenge where significant enough that it was definitely worth the effort and the results will most likely continue to increase the more data we are processing. 
