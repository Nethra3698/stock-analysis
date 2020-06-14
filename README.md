# Stock Analysis
Using VBA to analyse stock

## Challenge
The All Stock Analysis sub routine runs an analysis for the stocks based on the year that the user inputs. For the particular problem, the user has an option of inputting either 2017 or 2018 because of the availabilty of data but this code can be used for other years if data was added for the other years to the same workbook.
All the stocks are categorized under three headings, tickers, the total daily volume and the Return. Also depending on the percentage of the return, the stocks are color coded to indicate a drop or rise where a drop is indicated by red and a rise is indicated by green. 
The buttons can be used to clear the worksheet and run a new analysis. 

For the Challenge, the code was refactored so that it is much more flexible and efficient. One of the key points in the new edit is the use of arrays to hold the data. Previously, a single variable was used to hold a certain type of data and was updated to hold something else after outputting it on the worksheet. Now, this is not the case. Enough memory is allocated and nothing is erased so that if needed they can be used later. Also, in the new code the ticker is directly taken from the worksheets and not hardwritten into the code like it was previously. This allows for future changes and accomodates data of any stock and any year. 
