# stock-analysis
# Module 2- Visual Basic for Applications, a Dive into the Stock Market


## Overview of the Project 

### Purpose
##### For this assignment, we were presented with a simple question: “Should our client’s parents invest in Daqo, a New Energy Corporation?”. In order to answer this questions successfully, we were presented with a stock market data set for the years 2017 and 2018. From this data, we were able to utilize VBA to compute the the total daily volume, or the shares throughout the day, of each stock and discover the percentage increase or decrease in their yearly return. After initially running the analysis on Daqo, we expanded on our code to analyze and compare other green energy stocks to see if Daqo is the best investment choice.  In order to do this, we successfully applied our coding skills through the use of for loops and conditionals to direct logic flow, solve the problem at hand, and make the code interactive. After creating our initial code, it was further restructured and refactored to make the code more efficient and use less memory and potentially expand the data to include more stocks. 


## Analysis of Project 

### Analysis 
##### The stock dataset contains multiple green energy stocks along with their opening price, highest and lowest prices, their corresponding closing and adjusted closing price, and volume of shares throughout the day. For this analysis, we focused on the total daily volume for each stock and their closing price. In order to begin the analysis, we compiled an array of the stock tickers- which for our analysis is composed of 12 green energy stocks and is displayed in the image below. 

<img width="420" alt="TickerArray" src="https://user-images.githubusercontent.com/92558842/139619040-10fc6f77-98cb-4c89-8d36-f3a47d185c20.png">

We referenced this array throughout our code to analyze each stock individually. Furthermore, we refactored our code slightly to have the capability to accept more data values, thus more stocks, through the code below.

	'RowCount = Cells(Rows.Count, "A").End(xlUp).Row
	
This line is set to find the last value in our data set, giving us the number of cells to loop through in our analysis. With the number of cells in mind, we also looped through the array of tickers to compile data values for each ticker index. For instance, to find the total volume for each stock, we utilized a nested for loop where in the ticker volume increased by each value in the volume column for that ticker index and this repeated until data was collected for each ticker in the array. Also within this nested for loop, we implemented conditionals to find the first price land last price listed for each individual stock and continued this until we had the data for each stock. We later referenced the starting and ending prices to find the yearly return for each stock. The internal component of this nested for loop is shown the code below. 

<img width="575" alt="NestedForLoops_Conditionals" src="https://user-images.githubusercontent.com/92558842/139619126-d5901f01-a593-4e9b-a95b-311f87f68d05.png">

Still within the loop of the ticker array, we output our data for each ticker index by displaying the ticker, its accumulated volume, and its yearly percentage return, respectively, as shown in the code below. 

<img width="475" alt="DisplayingData" src="https://user-images.githubusercontent.com/92558842/139619159-4a412f70-083b-41fc-87ef-883ca19dff65.png">

We further restructured our code by making it more user friendly. We introduced formatting to aid in the readability of the data results and created options for user interaction through the use of message boxes and buttons. For example, the code below allows us to run the analysis on the year selected by the user: 

	'yearValue = InputBox("What year would you like to run the analysis on?")
	
Lastly, to properly refractor our code, we evaluated the run time of our code by using the timer function of VBA to measure how much time elapsed from the moment we receive the users input to the last value being output and formatted, as represented in the summary code below. 

	‘receive user input on what year to analyze
	startTime = Timer ‘timer begins
	‘code is ran and data accumulated
	‘data is output 
	‘data is formatted 
	endTime = Timer 
	MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue) ‘message displayed to user showing total run time 

### Results
##### The use of VBA allowed us to efficiently analyze multiple stocks with a click of a button. Through the code we were able to visualize a detailed summary of not only Daqo but an array of green energy stocks. This data is summarized for 2017 and 2018 below. 

#### Stock Analysis of 2017

<img width="324" alt="VBA_ChallengeResults_2017" src="https://user-images.githubusercontent.com/92558842/139619248-e961159e-009d-4de0-b470-70f6f7497e13.png">

![DailyVolume2017](https://user-images.githubusercontent.com/92558842/139619252-26863d02-9cf5-410d-9506-c387d6195681.png)

![YearlyReturn2017](https://user-images.githubusercontent.com/92558842/139619259-3345ea1a-2cb9-4d9b-aba6-b209f31d9833.png)

####Stock Analysis of 2018

<img width="323" alt="VBA_ChallengeResults_2018" src="https://user-images.githubusercontent.com/92558842/139619318-8fed8e47-0130-4b79-8692-642b18273164.png">

![DailyVolume2018](https://user-images.githubusercontent.com/92558842/139619325-eaaa9549-1e09-4fa9-b53c-3f9fca411078.png)

![YearlyReturn2018](https://user-images.githubusercontent.com/92558842/139619334-eb9c6d36-ad90-4679-8b8d-5be7e0bfacd9.png)


##### From this analysis we can gather a few things: the most obvious being that 2017 was a far better year than 2018. However, this was not secluded to green energy alone, rather it was correlated to the [stock market crash of 2018](https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018) where stocks plummeted as a result of slow economic growth and a fear of rising interest rates along with other factors. In regards to Daqo, we can see in 2017 that they were the leading in green energy for yearly return; however, their daily shares were not as reflective of their success as they continue to be low in both 2017 and 2018. Drawing on these results, I would be cautious to recommend Daqo to the client and would suggest they switch their investment to Enphase Energy which had positive yearly return for both 2017 and 2018 and continue to have high daily volume. 

#### In regards to refactoring, our code ran within hundredths of a second for both 2017 and 2018- improving by nearly ten-fold for each from our original code. The exact run times for our refactored code are displayed below, representing the effectiveness of refactoring code. 

<img width="264" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92558842/139619487-1f8fd857-d8b7-4335-beda-ce63d3cd58c5.png">

<img width="265" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92558842/139619932-20af6572-6a1f-48ab-a6a9-fcc30d388ba4.png">


## Summary

### Conclusion 
##### In conclusion, we discovered the effectiveness of refactoring by altering our original data so that it loops through the array of stocks one time to determine each tickers total volume and yearly percentage return. We were able to create a VBA macro that utilizes for loops and conditionals all while taking in user inputs, displaying pop ups, and allowing for changes in data. From these efforts, we can easily compare 2017 and 2018 green energy stocks and could potentially increase the number of stocks and, or add more years to analyze. We can also see from the data we do have, that Daqo may not be the best investment choice and can clearly see that 2018 was a rough year for the stock market. 

### Advantages and Disadvantages 
##### In general, our goal with coding is to 1. Make it run, 2. Make it better, and 3. Make it faster; thus, indicating how essential refactoring is to the coding process. The main advantages of refactoring are that it makes the code more efficient. This is done by recognizing patterns and taking fewer steps which will lead us to the next few advantages which are that refactoring improves logic and in turn, uses less memory. However, with every advantage there are disadvantages. What I have discovered is that refactoring takes time and genuine knowledge of the code that has been written- if one does not fully understand the code, it is difficult to refactor. Furthermore, there is an increase in the potential for errors while refactoring. For instance, one can copy and paste their original code and forget to change the variable or can lose track of where to place the original code into the refactored code. An advantage and disadvantage of refactoring is debugging- debugging helps improve the code in the long run but can be difficult to diagnose. 
##### Looking at our challenge, I found that the explicit lay out of our original code was advantageous for a small data set and aided in creating initial for loops and helped my general understanding of the problem at hand. The original code allowed us to create basic code fragments that can be reused not only for this project but for future challenges. The disadvantage was that it took a few seconds for the code to run to completion. It also had multiple macros which became difficult to keep track of. On the other hand, the advantages of the refactored code are that the program ran faster due to it having fewer steps and macros. For instance, instead of having a formatting macro like we did in the original code, we added the formatting to the loop through tickers so it was completed as the analysis was ran. A personal disadvantage for the refactored code was that it took longer than expected to debug while combining the for loops from the original code into a nested for loop and I struggled with an overflow error. A disadvantage of both the refactored and original code is that the array of tickers is limited to what we have input manually. If the client wanted to add more stocks to analyze, we would have to go into VBA and manually input that ticker into the array and alter the size of the array to loop through in order for the data to work properly. In essence, the advantages of the refactored code and the end result far outweigh the original. 

