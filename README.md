# Stock Market Data Analysis with VBA

## Background
In this homework assignment, I used VBA scripting to analyze generated stock market data. The goal was to loop through the stocks for one year, calculate various metrics, and identify stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. Conditional formatting was applied to highlight positive and negative changes in green and red, respectively.

## Files
The project files can be accessed using the following link: VBA Challenge Files

## Analysis Instructions
To analyze the stock market data, I performed the following tasks using VBA scripting:

* Loop through Stocks: I created a script that loops through all the stocks for one year and extracts the required information for each stock.

* Yearly Change: I calculated the yearly change for each stock by subtracting the opening price at the beginning of the year from the closing price at the end of the year.

* Percentage Change: I calculated the percentage change for each stock by dividing the yearly change by the opening price and multiplying by 100.

* Total Stock Volume: I calculated the total stock volume for each stock, which represents the cumulative volume of shares traded throughout the year.

* Conditional Formatting: I applied conditional formatting to highlight positive changes in green and negative changes in red, making it easier to identify performance trends.

* Greatest Metrics: I added functionality to the script to identify the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Other Considerations
To enhance the functionality and efficiency of the VBA script, the following considerations were made:

* Scalability: The script was adjusted to run on every worksheet (representing each year) in the workbook. This enables the analysis to be performed on multiple years' worth of stock market data with a single click.

* Dataset Selection: During the development and testing of the code, the "alphabetical_testing.xlsx" dataset was used. This smaller dataset allows for quicker testing and validation of the script, ensuring it runs within a reasonable time frame (under 3 to 5 minutes).

Please refer to the VBA script files included in the project repository for a detailed view of the code implementation and data analysis.

Note: This README file provides an overview of the project and the actions taken. For more specific details, please refer to the project files and the VBA script itself included in the shared link.
