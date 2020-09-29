# Green Stock-analysis

## Overview of Project: 

The purpose of this analysis is to offer a clear and simple way for our friend Steve to quickly analyse stock data for different years and to do it in as efficient a manner as possible. In this case when we refactored the code we removed the nested for loop.

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.


Our analysis has shown that with the exception of TERP stock all of these stocks did not do as well in 2018 as they did in 2017. The stocks ENPH & RUN both still had positive growth in 2018 so those would be the best stock to invest in for the long term. Also the stock DQ that Steve's parents had previously invested in had the highest returns in 2017 but the most losses on 2018 so they are very volatile and not good for long term investments. 

The refactored code was significantly faster to run. 

Unrefactored results for 2017 returned in `1.3359938` seconds while the refactored code took `0.3125` seconds. For the year 2018 the unrefactored code took `1.429688` seconds and the refactored code returned results in `.021875` seconds. Looping through the main data one but introducing a separate For loop to move through the tickers made a significant improvement in the time it took to run the code. 

###2017 Results with Refactored Code:

![Refactored Results & Run Time for 2017](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

###2018 Results with Refactored Code:

![Refactored Results & Run Time for 2018](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

###2017 Results with Unrefactored Code:

![Un-Refactored Results & Run Time for 2017](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2017_unrefactored.png)

###2017 Results with Unrefactored Code:

![Un-Refactored Results & Run Time for 2018](https://github.com/ccastanette/stock-analysis/blob/master/Resources/VBA_Challenge_2018_unrefactored.png)


## Summary: 

### What are the advantages or disadvantages of refactoring code?

### How do these pros and cons apply to refactoring the original VBA script?
