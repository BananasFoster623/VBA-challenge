# VBA-challenge
Homework 2 for the Georgia Tech Professional Education Data Science and Analytics BootCamp. This project took multiple stock's price data over time and extracted the following:
- Annual change in trading price for each stock
- Percent change in trading price for each stock
- Total trading volume for that calendar year for each stock

There were additional, optional data points that were requested. They are:
- Stock with the greatest percent increase in trading price for each calendar year
- Stock with the greatest percent decrease in trading price for each calendar year
- Stock with the greatest trading volume for each calendar year

# Summary of Deliverables
This repository contains the following:
- The Visual Basic Script (.vbs) file containing the VBA code used for the analysis
- Screenshots of the resulting Excel Worksheets verifying the code executed correctly
- This detailed readme file

# Code Description
The VBA code contains five (5) sub procedures named (listed in order they are called):
- Main
- SetHeaders
- Work
- PrintOut
- Formatting

Variables were declared at the module level, outside of any sub procedures. 
  
### Main
This is the sub procedure that should be called to exectute the code. It loops over each Worksheet in the Excel Workbook to analyze each year of data with a single macro excution. When successfully completed, a message box will alert the user that the code was executed successfully with no errors. 
 
### SetHeaders
This sub procedure simply sets the column headers for the desired outputs, which were defined in the problem statement. 

### Work
This is the sub procedure that does all of the data manipulation and calculations. In order of execution, the commands can be summarized as follows:
1. Finds the length of the dataset on the active worksheet
2. Finds the number of unique ticker symbols (column 1) for use in loops
3. Redim
