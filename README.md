# Problem Statement

The current process for this report relies heavily on Microsoft Excel spreadsheet manipulation. I run extracts from Salesforce for 12 months, running one month at a time. Once that is complete, then we have to go through each month and manipulate some of the data to clean it and to correct any errors. Afterwards, all the months are put into one tab in a Microsoft Excel workbook.

# Solution

Use python to open each file in a folder with the downloads and then manipulate and clean the data using the pandas library. 

Once the cleaning is complete, then merge all the dataframes into one dataframe and export to a Microsoft Excel workbook.