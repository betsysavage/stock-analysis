# Stock Analysis
## Overview of Project ##
## Purpose ##
In this challenge, our client, Steve, requested some help analyzing stock performance in order to advise his parents about investment strategies. This analysis involved setting up a macro to automate calculations of various measures of stock performance, including the daily volume of stocks traded throughout the day and the yearly return in price, for the years 2017 and 2018. Using the Visual Basic functionality of Excel allows us to apply formulas to several stocks without individually manipulating calculations for each, and to repeat these formulas in a "loop" pattern over a specified number of values. While writing these macros enabled us to complete the stock analysis more efficiently, the way code is written can have an impact on its speed as well, so this analysis also included refactoring (editing) the available code with the goal of making it run as quickly as possible. Maximizing the efficiency of the script will allow Steve the option to apply this analytical tool to a broader number of stocks.  
-
### Results ###
**1. Compare the stock performance between 2017 and 2018**

The difference between the two years is quite large. While 11 out of 12 stocks were traded with positive returns in 2017, 2018 saw an overall decline in stock performance. 


<img width="469" alt="image" src="https://user-images.githubusercontent.com/114873837/202518621-fefc80d4-42af-467a-8796-895bb0a4b2e4.png">



This trend of strong 2017 performance, followed by declining 2018 stock values, seems to be true for the majority of stocks listed in our analysis. The only exceptions are companies with the ticker names of "ENPH" and "RUN", which perform strongly with positive returns in both years. The only company that is unsuccessful in both years is "TERP". It is recommended that Steve conduct further analysis to explain this overall trend so he can best advise his parents. Are there common traits in the 12 companies examined within this analysis (i.e. same industry, similar products, or other market forces) that could explain this downward trend? Are there common traits or strategies used by ENPH and RUN that would have allowed them to succeed in a year when other stock values struggled? Did the stock prices and trading patterns recover the following year? Answering these questions will allow Steve to hypothesize how results of this analysis would appear across a larger sample size, and would better equip Steve to forecast stock performance and advise his parents' investment strategy.  

**2. Compare the execution times of the original script and the refactored script**

The original script produced a result in 0.078125 seconds for the 2017 stock analysis, and a result in 0.08984375 seconds for the 2018 stock analysis.

<img width="420" alt="image" src="https://user-images.githubusercontent.com/114873837/202326334-06ce224c-ff92-42fc-82a0-09101c5a9e77.png">
<img width="431" alt="image" src="https://user-images.githubusercontent.com/114873837/202326852-4f67be95-98ab-42c1-b2e1-761cc89e28e0.png">

The refactored script produced a result in 0.046875 seconds for the 2017 stock analysis, and a result in 0.046875 seconds for the 2018 stock analysis.

<img width="265" alt="image" src="https://user-images.githubusercontent.com/114873837/202326659-84a01639-e214-411b-9249-d94e4b7c7525.png">
<img width="265" alt="image" src="https://user-images.githubusercontent.com/114873837/202326726-be73692a-1c0e-4850-a2a3-8f1f67564820.png">

While the speed of these results is fast for both versions of the script, the time being saved in the refactored script adds up. Refactoring the script has resulted in the run time being reduced by nearly 50%! This workbook only focuses on 12 stock tickers, but ultimately Steve wants to expand his analysis to the entire stock market. When examining such a large data set, the efficiency of the script will be crucial in effectively generating results.

The refactored code can work more quickly because the for loop has been set up to loop through the ticker index numbers, of which there are 12, instead of each individual row.

### Summary ###
1. What are the advantages and disadvantages of refactoring code?
Refactoring code makes it easier for users to understand and make updates. It requires adding comments and arranging lines of code to make the purpose as clear and efficient as possible. The process of refactoring can also uncover bugs that may have been interfering with functionality. While it takes time, investing in refactoring can allow code to be "cleaned up" so it can eventually be applied on a larger scale.
The primary disadvantage of refactoring code is its cost, both in time and money, with no real change in functionality of the output. Rearranging code to make it more readable and efficient can take a good deal of staff time, between debugging and reviewing. I struggled at several points in this analysis with error messages that prevented the script from running at all, which made it challenging to test different pieces of the code. It is tempting to fall back on the available option that works, rather than committing to a time confusing and sometimes frustrating process of editing. Ultimately, however, I broke the code into shorter blocks to test each segment, and the ultimate outcome of 
(Background research for this question was found on https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)

2. How do these pros and cons apply to refactoring the original VBA script?
