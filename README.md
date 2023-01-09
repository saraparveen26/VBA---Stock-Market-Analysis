VBA Challenge – an analysis of Stock Market data

In the second challenge of Data Analytics Bootcamp, I have analyzed data provided on Stock Market using Visual Basic. We were provided with two files: one for testing the code on a smaller data size, and the second containing full data.
The testing file contained 6 sheets named A through F. The actual bigger file contained 3 sheets containing data for the years 2018, 2019 and 2020.
I tested my data on both the files and it ran successfully without any errors and the output results were displayed in all sheets on each file. The screenshots submitted show results for each file and each sheet on the respective file. I have captured the screenshots displaying top and bottom of the output results.


The code starts by setting the loop to run through all worksheets and clearing any previous formatting.
In the next steps, the number of filled rows are counted and then headers are added and formatted to be bold for columns I through L to display outputs for Ticker, Yearly Change, Percent Change, and Total Stock Volume.
Then variables are declared to navigate down the input and output columns as well as storing values for open price, close price, yearly change, percent change and total stock volume.

Unique Ticker/ Stock Values:

A loop is initiated to navigate through the original stock data and display unique ticker/ stock symbols into column I with header “Ticker”. Starting value of 2 is assigned for the variables navigating through column I in order to exclude the header row. The loop navigates down the output column for each unique value display.

Yearly Change, Percent Change and Total Stock Volume:

Starting value of 2 is assigned for the variables navigating through different columns in order to exclude the header row. A starting value of zero is started for Total volume. A loop is initiated to navigate through the original stock data.  If statements are used to retrieve open and close prices for each unique ticker symbol which in turn are used to calculate yearly change and percent change. Volume from column G is added on a cumulative basis for each unique ticker symbol and resets to zero once a new ticker symbol starts in the original data set. The loop navigates down the output columns of Yearly Change, Percent Change, and Total Stock Volume as well as the input data for open price, close price and volume data.

Conditional Formatting:

Column J and K displaying Yearly Change and Percent Change respectively have been formatted conditionally using Visual Basic Code. The cell fill color changes to Red if the cell value is less than zero, it changes to Green if the cell value is greater than zero, and it changes to Yellow if the cell value is equal to zero.

Greatest Increase, Decrease and Total Volume:

Headers are added for columns and rows to display the values for greatest % increase, greatest % decrease, greatest total volume and the respective ticker symbols and are formatted to be bold. Then the number of filled rows are counted for the previously retrieved output data of unique ticker values and variables are declared to store greatest increase, decrease and stock volume.  Starting value of 2 is assigned for the variables navigating through columns I through K in order to exclude the header row. Starting values of zero are assigned to the variables storing the greatest Increase, Decrease and Stock Volume so that it is reset as the loop runs through the next Worksheet. 

A loop is then started to run through the data extracted for unique ticker values and their respective yearly change and volume. If statements are used to retrieve and store the values for the greatest increase, decrease and volume as well as output the respective ticker symbols in column P and Q.
I am submitting the solution using loops to calculate and retrieve the greatest values. Alternatively, these can be retrieved using Worksheet Functions of Min and Max. I had tried the code with these functions and it presented the same results.

Lastly, AutoFit is applied to adjust the column width so that values are displayed properly.
The loop then runs through the next Worksheet once all the above steps are completed.

This analysis presents a meaningful overview of data and provides the basis for comparing different stocks for their increase/decrease and trading volume over the year. The greatest numbers identify the stocks with significant movements over the year in terms of price changes and volumes and can be very helpful in exploring further insights about these stocks.
