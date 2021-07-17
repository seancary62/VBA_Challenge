# VBA_Challenge
A stock market analysis using VBA

#Stock Market Analysis Using VBA

##Overview
For this project we were tasked with assisting our client, Steve, and his parents in developing a VBA script that could efficiently analyze stock market data for multiple years. In the first iteration, we created a script that could easily process 12 stock tickers for a selected year. However, the client is now requesting that the script be refactored in order to handle the entire stock market.

###Purpose
The purpose of this project is to refactor the provided VBA script in order to make it more efficient. This will make it easier to work with larger sets of data.

##Analysis
In order to refactor our data we created 3 additional arrays to store the trading volume, starting price, and ending price for each ticker. A nested for loop was used to cycle through each ticker in the index. This was done in order to sum the trading volume and identify the starting and ending prices for the year selected by the user. This information was then stored in the arrays with the corresponding index number. Another loop was used to print the information on the spreadsheet for each ticker. The timer function in VBA was used to record the amount of time the script took to process.

##Results
Using this method resulted in a faster processing time for the output as oppposed to the other method of looping through and printing at each iteration. Using the refactored version of the script we saw times around 0.53s to process the 12 stock tickers in the spreadsheet versus the previous method which took about 0.75s. For 12 stock tickers this may not seem like a huge improvement, but if we consider using the script for hundreds or thousands of stock tickers it should have a significant impact in the performance capabilities of the subroutine.

##Summary and Conclusions
From the results it can be stated that refactoring data is an important process in coding, as it can result in significant improvements in performance and expand the capabilities of a script. By refactoring  the data using arrays we were able to improve the efficiency of the VBA script so that Steve can analyze a larger number of stocks for his purposes.
