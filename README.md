# Alternative Green Energy Stock Analysis Using Excel VBA

##Overview Of Project 

  For this project, I was expected to use Excel VBA as a tool to analyze the performance of multiple green energy stocks, for the years 2017 and 2018. By using Excel VBA, I created a macro that allowed me to run both sets of data, in a short amount of time, and created a formatted, easy to read table. The table was created to make it easy for a potential investor, who is trying to decide which stocks were traded in higher volume, and which stocks produced a higher return. Although my macro ran fine initially, it was then expected that I try to find a way to make my macro run even faster or more efficiently. 
	

###Purpose

  The purpose of the analysis was to look at green energy stock companies, for the years 2017 and 2018, and to determine which forms of green energy is best to invest into. Green energy stocks were looked at, due to the belief that as fossil fuels get used up, there will be more reliance on alternative energy production in the future. The analysis was performed on 12 different green energy stock companies, and I was given thousands of cells of data, which would take forever to manually count up or produce a table to sum up each stock’s total volume and return percentage per year. Thus, by using excel VBA, I could produce the same table that could run this analysis for either year, depending on what value was entered into an initial pop-up box, prior to running the macro. 
    The initial macro created ran quickly, but there was only 12 different stock company data to run through, so if there were thousands of different stock companies to analyze, it may not be as quick. This is why refactoring the code with less steps, in the scripting process, which would use less computer memory, and could make it easier for someone else in the future to read and possibly use for their own data analysis projects.  Code can be written in so many different ways, that one way might not always be the easiest or most efficient way to produce the same result, which is why refactoring is an integral part to the coding process. 

### The Data 

	First, I created a macro, “DQ Analysis”, that ran the total daily volume, and return for one stock company, “DQ”, which I used as a base to build on for when I wanted to create a new macro to analyze all 12 stock companies. The second macro, “All Stocks Analysis”, started out similarly to the first one, by formatting the output sheet and setting up variables so my for-loops would run correctly, but one major difference was having to initialize an array of all my stock tickers. Setting up a ticker array allowed for my macro to loop through all the stocks, instead of just one. 
	Loops tell a computer to run a code repeatedly, and for a specific number of times, because it uses an iterator. An iterator is a named variable that will change its value with each run of the for-loop, increasing by one. I set up the starting and ending price variables, which I used in my for-loops. I created two for-loops, an outside loop so that VBA would loop through all the tickers, and another nested inside the outside loop, which had if-then statements. One if-then statement calculated the total volume of stocks traded for each ticker, and then two other if-then statements, that calculated a starting price and an ending price for each ticker. I had to set the totalVolume variable to zero, which allowed me to do two things, add to the totalVolume variable inside of the loop, as well as giving me a sum of the total volume for each ticker. Following the for-loops, I created outputs in my “All Stocks Analysis” worksheet for where the tickers, their total volumes, and their total return values would be presented. 
	Finally, I created a small macro that would format the table headers, and color the interior of the total return values, as well as formatting the number values. The total return values box would be the color green if the return resulted in a positive number, or above 0%, but if the return value was a negative number, or less than 0%, then the interior color would be red. 
	For the refactored code, all of the previous steps from the “AllStocksAnalysis” macro were relatively the same, except a couple of steps. For instance, instead of creating nested loops, four for-loops were created, and another variable was added, named the “tickerIndex”. This variable was used in all three for-loops. In the first one, it was used to initialize the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to zero. In the second for-loop, it was used in an If-then statement to increase the tickerVolumes for the current ticker being ran at that position in the loop. In the third loop, it was used in another if-then statement to calculate the ticker starting and ending prices. A final for loop was created to loop through the ticker arrays and output those values into a table, under the column headers, “Ticker”, “Total Daily Volume”, and “Return”. 
	An InputBox() command was used on both macros, in which, this command works like a message box, asking the user which year they want to run the analysis on, and the user can input and choose which year’s data that they want to run and analyze. An example of the inputBox() command prompt is shown below:
  
  ![Input_box_command_prompt](https://user-images.githubusercontent.com/104864579/174910499-6e041bd6-9ecf-4c0d-b3f1-ba300db91e74.png)


## Results 
### Analysis 
  Based on the stock performance data table for the year 2017, almost all of the stock tickers showed a positive return, except “TERP”, which was the only ticker in the red. 
  ![stocks_analysis_2017](https://user-images.githubusercontent.com/104864579/174910887-1e73fade-7fb2-48b6-ba45-7911ea33208f.png)

  When I ran the macro a second time, inputting 2018 into the input box, almost all of the stock tickers were in the red, except two tickers, “ENPH” and “RUN”. 
  ![stocks_analysis_2018](https://user-images.githubusercontent.com/104864579/174910924-4564abb4-c40f-4b26-930b-0fa462cfbb92.png)

  By looking at the return data values for 2017 and 2018, it can be determined that the stock performance for “ENPH” and “RUN” showed positive returns for both years, and effectively being better investments comparatively to the other 10 stocks. 
  
  In order to determine which code ran quicker or more efficiently, a Timer function was used. Two variables, startTime and endTime, were initialized and each set equal to the Timer function. The startTime variable was placed at the beginning of the subroutine, and the endTime variable was placed at the end of the subroutine. These variables were purposely placed after the InputBox() command, so that the timer would start only when the macro starts running, and not include the time it takes a user to input which year they want to analyze. 
    
