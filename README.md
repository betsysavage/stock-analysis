# Stock Analysis#
## Overview of Project ##
## Purpose ##
In this challenge, our client, Steve, requested some help analyzing stock performance in order to advise his parents about investment strategies. This analysis involved setting up a macro to automate calculations of various measures of stock performance for the years 2017 and 2018. Using the Visual Basic functionality of Excel allows us to apply formulas to several stocks without individually manipulating calculations for each, and to repeat these formulas in a "loop" pattern over a specified number of values. While writing these macros enabled us to complete the stock analysis more efficiently, the way code is written can have an impact on its speed as well, so this analysis also included refactoring (editing) the available code with the goal of making it run as quickly as possible. Maximizing the efficiency of the script will allow Steve the option to apply this analytical tool to a broader number of stocks.  
-
### Results ###
**1. Compare the stock performance between 2017 and 2018**
The difference between the two years is quite large. While 11 out of 12 stocks were traded with positive returns in 2017, 2018 saw an overall decline in stock performance, with all but two companies indicating negative returns for the year. 

**2. Compare the execution times of the original script and the refactored script**

The original script produced a result in 0.078125 seconds for the 2017 stock analysis, and a result in 0.08984375 seconds for the 2018 stock analysis.

<img width="420" alt="image" src="https://user-images.githubusercontent.com/114873837/202326334-06ce224c-ff92-42fc-82a0-09101c5a9e77.png">
<img width="431" alt="image" src="https://user-images.githubusercontent.com/114873837/202326852-4f67be95-98ab-42c1-b2e1-761cc89e28e0.png">

The refactored script produced a result in 0.04687 seconds for the 2017 stock analysis, and a result in 0.046875 seconds for the 2018 stock analysis.

<img width="265" alt="image" src="https://user-images.githubusercontent.com/114873837/202326659-84a01639-e214-411b-9249-d94e4b7c7525.png">
<img width="265" alt="image" src="https://user-images.githubusercontent.com/114873837/202326726-be73692a-1c0e-4850-a2a3-8f1f67564820.png">

While the speed of these results is fast for both versions of the script, the time savings add up. Refactoring the script has resulted in the run time being reduced by nearly 50%! This workbook only focuses on 12 stock tickers, but utlimately Steve wants to expand his analysis to the entire stock market. When examining such a large data set, the efficiency of the script will be crucial in effectively generating results.


### Summary ###
1. What are the advantages and disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
