# Alternative Green Energy Stock Analysis Using Excel VBA

## Overview Of Project 

  For this project, I was expected to use Excel VBA as a tool to analyze the performance of multiple green energy stocks, for the years 2017 and 2018. By using Excel VBA, I created a macro that allowed me to run both sets of data, in a short amount of time, and created a formatted, easy to read table. The table was created to make it easy for a potential investor, who is trying to decide which stocks were traded in higher volume, and which stocks produced a higher return. Although my macro ran fine initially, it was then expected that I try to find a way to make my macro run even faster or more efficiently. 
	

## Purpose

  The purpose of the analysis was to look at green energy stock companies, for the years 2017 and 2018, and to determine which forms of green energy is best to invest into. Green energy stocks were looked at, due to the belief that as fossil fuels get used up, there will be more reliance on alternative energy production in the future. The analysis was performed on 12 different green energy stock companies, and I was given thousands of cells of data, which would take forever to manually count up or produce a table to sum up each stock’s total volume and return percentage per year. Thus, by using excel VBA, I could produce the same table that could run this analysis for either year, depending on what value was entered into an initial pop-up box, prior to running the macro. 
  
  The initial macro created ran quickly, but there was only 12 different stock company data to run through, so if there were thousands of different stock companies to analyze, it may not be as quick. This is why refactoring the code with less steps, in the scripting process, which would use less computer memory, and could make it easier for someone else in the future to read and possibly use for their own data analysis projects.  Code can be written in so many different ways, that one way might not always be the easiest or most efficient way to produce the same result, which is why refactoring is an integral part to the coding process. 

## The Data 

  First, I created a macro, “DQ Analysis”, that ran the total daily volume, and return for one stock company, “DQ”, which I used as a base to build on for when I wanted to create a new macro to analyze all 12 stock companies. The second macro, “All Stocks Analysis”, started out similarly to the first one, by formatting the output sheet and setting up variables so my for-loops would run correctly, but one major difference was having to initialize an array of all my stock tickers. Setting up a ticker array allowed for my macro to loop through all the stocks, instead of just one. 
  
    Sub AllStocksAnalysis()

      Dim startTime As Single
      Dim endTime As Single

      yearValue = InputBox("What year would you like to run the analysis on?")

      startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet

      Worksheets("All Stocks Analysis").Activate
    
      Range("A1").Value = "All Stocks(" + yearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
    '2) Initialize array of all tickers

     Dim tickers(11) As String
    
      tickers(0) = "AY"
      tickers(1) = "CSIQ"
      tickers(2) = "DQ"
      tickers(3) = "ENPH"
      tickers(4) = "FSLR"
      tickers(5) = "HASI"
      tickers(6) = "JKS"
      tickers(7) = "RUN"
      tickers(8) = "SEDG"
      tickers(9) = "SPWR"
      tickers(10) = "TERP"
      tickers(11) = "VSLR"

  Loops tell a computer to run a code repeatedly, and for a specific number of times, because it uses an iterator. An iterator is a named variable that will change its value with each run of the for-loop, increasing by one. I set up the starting and ending price variables, which I used in my for-loops. I created two for-loops, an outside loop so that VBA would loop through all the tickers, and another nested inside the outside loop, which had if-then statements. One if-then statement calculated the total volume of stocks traded for each ticker, and then two other if-then statements, that calculated a starting price and an ending price for each ticker. I had to set the totalVolume variable to zero, which allowed me to do two things, add to the totalVolume variable inside of the loop, as well as giving me a sum of the total volume for each ticker. Following the for-loops, I created outputs in my “All Stocks Analysis” worksheet for where the tickers, their total volumes, and their total return values would be presented. 
  
    '3a) Initialize variables for starting price and ending price

    Dim startingPrice As Single
    Dim endingPrice As Single

    '3b) Activate data worksheet

    Worksheets(yearValue).Activate
    
    '3c) Get the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers

    For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0
   
    '5) loop through rows in the data

       Worksheets(yearValue).Activate
        For j = 2 To RowCount
    
       '5a) Find total volume for current ticker
         If Cells(j, 1).Value = ticker Then
         
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
       '5b) get starting price for current ticker
           
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            startingPrice = Cells(j, 6).Value
        
        End If
           
        '5c) get ending price for current ticker
           
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            endingPrice = Cells(j, 6).Value
            
        End If
    
    	Next j
        
	'6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i

  Finally, I had to format the table headers, assign a color to the interior, or background, of the total return values, as well as formatting the number values. The total return values box would be the color green if the return resulted in a positive number, or above 0%, but if the return value was a negative number, or less than 0%, then the interior color would be red. 
  
	Range("B4:B15").NumberFormat = "#,##0"
	Range("c4:c15").NumberFormat = "0.00%"
	
	dataRowStart = 4
	dataRowEnd = 15

	 For i = dataRowStart To dataRowEnd

    	If Cells(i, 3) > 0 Then
        'color the cell green
    
       	Cells(i, 3).Interior.Color = vbGreen
    
   	ElseIf Cells(i, 3) < 0 Then
       	'color the cell red
    
       	 Cells(i, 3).Interior.Color = vbRed
    
    	Else
       	'clear the cell color
    
        Cells(i, 3).Interior.Color = xlNone
    
  	End If

	Next i

## The Data Refactored

  For the refactored code, all of the previous steps from the “AllStocksAnalysis” macro were relatively the same, except a couple of steps. For instance, instead of creating nested loops, four for-loops were created, and another variable was added, named the “tickerIndex”. This variable was used in all three for-loops. In the first one, it was used to initialize the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to zero. In the second for-loop, it was used in an If-then statement to increase the tickerVolumes for the current ticker being ran at that position in the loop. In the third loop, it was used in another if-then statement to calculate the ticker starting and ending prices. A final for loop was created to loop through the ticker arrays and output those values into a table, under the column headers, “Ticker”, “Total Daily Volume”, and “Return”. 
  
    '1a) Create a ticker Index
    ticker = tickers(i)
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                       
            End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
  
### User Interaction  

  An InputBox() command was used on both macros, in which, this command works like a message box, asking the user which year they want to run the analysis on, and the user can input and choose which year’s data that they want to run and analyze. An example of the inputBox() command prompt is shown below:

  
  ![Input_box_command_prompt](https://user-images.githubusercontent.com/104864579/174910499-6e041bd6-9ecf-4c0d-b3f1-ba300db91e74.png)


## Results 

### Analysis of Stock Performance
  Based on the stock performance data table for the year 2017, almost all of the stock tickers showed a positive return, except “TERP”, which was the only ticker in the red. 
  
  ![stocks_analysis_2017](https://user-images.githubusercontent.com/104864579/174910887-1e73fade-7fb2-48b6-ba45-7911ea33208f.png)

  When I ran the macro a second time, inputting 2018 into the input box, almost all of the stock tickers were in the red, except two tickers, “ENPH” and “RUN”. 
  
  ![stocks_analysis_2018](https://user-images.githubusercontent.com/104864579/174910924-4564abb4-c40f-4b26-930b-0fa462cfbb92.png)

  By looking at the return data values for 2017 and 2018, it can be determined that the stock performance for “ENPH” and “RUN” showed positive returns for both years, and effectively being better investments comparatively to the other 10 stocks. 
 
### Analysis Using Timer Function 
  In order to determine which code ran quicker or more efficiently, a Timer function was used. Two variables, startTime and endTime, were initialized and each set equal to the Timer function. The startTime variable was placed at the beginning of the subroutine, and the endTime variable was placed at the end of the subroutine. These variables were purposely placed after the InputBox() command, so that the timer would start only when the macro starts running, and not include the time it takes a user to input which year they want to analyze. 
    
  When I ran the Timer function for the original code, the message box for 2017 showed that the run time was 0.375 seconds, and the message box for 2018 showed that the run time was 0.3671875 seconds, as shown below: 
  
  ![Original_code_2017](https://user-images.githubusercontent.com/104864579/174913146-5b7f9b68-10e2-4798-8489-13957ab6f8b7.png) 
  ![Original_code_2018](https://user-images.githubusercontent.com/104864579/174913175-a092cbfc-220f-4e27-b81a-247c6248194a.png)

  When I ran the Timer function for the refactored code, the message box for 2017 showed that the run time was 0.05859375 seconds, and the message box for 2018 showed that the run time was 0.06640625 seconds, as shown below: 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/104864579/174913295-5bc92070-9c8e-4e46-8a15-0ef8a5149213.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/104864579/174913321-b72c8967-1b90-4859-9e9b-8950fc374450.png)

  Based on the run time for the original and refactored code, it can be concluded that the refactored code run time was significantly faster than the run times in the original code for the years 2017 and 2018. 

## Summary 

### Pros and Cons of Refactoring Code 

Refactoring code allows one to find multiple ways to try to solve a problem, while still producing the same end result. Based on the fact that there are multiple ways, it is not hard to understand that there are going to be ways that are better than others. Refactoring can help make our codes more organized, and easier to read for another person trying to understand or use that code for a project they might be working on. Obviously, they would have to edit certain variables or tickers for their specific for-loops, but if the code is easy to read and understand, then they can make quick changes and run the macro without producing an error. Refactoring code can also be helpful in software development and improvement, or make programming faster for products on the market that want to add more functions/features without disrupting the quality of the original program. A disadvantage of refactoring code could result in processes that might not be more efficient, or might not run well on other systems or applications that do not use the same test codes. 

### The Advantages of Refactoring the Original VBA Script

An advantage of refactoring the original VBA script, was that the macro run time decreased substantially. The benefit of this advantage would come to light if I was running a code that contained huge amounts of data or had to be ran for a much higher amount of tickers. This would also be an advantage if I created a macro that contained hundreds of lines of code. In this challenge, refactoring the code saved me 10ths of a second, but for a much bigger code or macro, the amount of time saved could accumulate into saving minutes to hours, or days even. Another advantage to refactoring the original VBA script would be that it allowed me to see that there are multiple ways that I can write a script and still get the same result.


