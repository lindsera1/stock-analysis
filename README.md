# Using VBA for Stock Analysis On Green Stocks

## Objective & Purpose

This project looked at the performance of 12 "Green Stocks" over the years 2017 and 2018. Specifically, we looked at the total
volume that a stock sold in shares, and the price of the stock at the beginning of the year against the end of the year, and 
calculated the returns from those two info points. By getting these performance metrics, we hoped to give Steve a more informed
view of which stocks to recommend for his parents, and which ones to avoid!

## Analysis

### DQ

The first stock of choice Steve's parents were keen of was the first one we analyzed. Using the parameters of volume and return,
we took a close look at how DQ performed in 2017 and 2018. 

![DQ_Results.png](https://github.com/lindsera1/stock-analysis/blob/master/DQ_Results.png)

Looking at 2018 performance, we might want to suggest DQ to Steve's parents. We used simple Excel macros in VBA to calculate 
the total volume as well as determing the starting price at the beginning of the year, and the ending price at the end of the
year. From these last two data points, we were able to calculate the total return, which turned out to be negative 63 percent. 

In light of how poorly DQ performed in 2018, we did not want to leave Steve's parents optionless and went the extra step of
evaluating 11 other Green stocks.

### The Other Green Stocks Analysis

Here, we evaluated the performance of 12 other green stocks in 2018 using the same parameters we used for the DQ Analysis.
This is what we found:

![All_stock_results.png](https://github.com/lindsera1/stock-analysis/blob/master/All_stock_results.png)

To make the evaluation simpler, we color coded the stocks that gave positive return percentages the color green, and those
returning negative percentages, the color red. Knowing this, we can see that two stocks seemed to perform better than all the
others ("ENPH" and "RUN").

When I ran a 2017 analysis, all but 1 stock yielded positive returns, but wanting to be as relevant to the current time as 
possible we should only base our recommendations on the most recent performance. In the future, if Steve wanted to look at more 
stocks, this could require a change in our approach in VBA to how we handle large amounts of data, so we refactored our code.

## Refactoring Code

### The General Pros and Cons

This process is always done with the intention of making current code better in some way, and with the knowledge that the first
shot at using code to solve problems, might not be the best way to do it. Especially if Steve wanted to run an analysis on many 
other stocks, we'd have to write our code in a way so as to avoid excess verbosity, eliminate reptitiveness, and keep our code
relatively short. To do this, we use a process called refactoring, which works somewhat like refinement. Refactoring hopefully 
makes our code easier to read and change in the future when other programmers look at it. On the downside of refactoring, 
unintended consequences may happen, as in my case, the run time of the code became longer.

### Specific Pros and Cons

I managed to narrow our loop down to one loop to perform all the assignments, rather than using a nested
loop.

![Refactored_Code2.png](https://github.com/lindsera1/stock-analysis/blob/master/Refactored_Code2.png)

By doing this, and placing the results of volumes, starting prices, and ending prices all in arrays, we allow the end user
to tack on as many tickers as he wishes, so long as they stay in order. Once we have looped through every row of data, we have
the volumes, starting prices, and ending prices of all the tickers. From these, we can use a simple loop to pass over 
every element in the array and output the desired information.

On the downside, this code actually took longer to run, and I could not determine why since every row was only getting looped 
over once, as opposed to 12 times. Perhaps there was more to evaluate per row, so it took longer. In spite of the 2 second
increase in run time, I would still say that the code is much more readable, and able to give future users the option of adding
more stocks to evaluate in the future. 
