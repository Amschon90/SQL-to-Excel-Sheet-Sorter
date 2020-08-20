# SQL-to-Excel-Sheet-Sorter
A script I wrote to sort an excel doc exported from DBVisualizer into a bunch of sheets.

## Background
It is expected you have a single sheet of data in .xlsx format exported from a database, with headers in row #1, and you want to sort this data into a bunch of different sheets based on whatever is in column #1. For example, if you had written a query like below:
  
  SELECT * FROM ORDERS
  ORDER BY CUSTOMER_ID, DATE_ORDERED;
  
 .. You would recieve a giant spreadsheet like this:
 
 CUSTOMER ID | ITEM | ORDER_DATE | COST
 001234567    PIZZA    7-4-2019   12.45
 001234567    CAKE     7-9-2019   19.89
 005665452    DONUT    6-1-2019   3.25
 081244241    DONUT    8-2-2019   3.25
 005665452    SODA     9-2-2019   1.99
 
But what if, instead of seeing everything in single sheet, you wanted it sorted into individual sheets for each customer, in a single Excel doc? That's what this script is for.

## How to Use
Ensure you have openpyxl installed, if not install it through command prompt:
  pip install openpyxl
  
Then just run the script, supply it a path and the file name you want sorted in that path (without the file extension), and it should complete in a matter of moments. Your sorted file will appear in the same path you specified earlier with a simular name, but worth "sorted" added to the file name.

Please note this script will not touch any other files in the folder.
