#### StockPriceUpdate

Proof of concept: Running the notebook will read in an excel-file, expected to have a column of stock-names on its first sheet. On a new sheet, StockPriceUpdate, for every name, it will get the most recent price of that stock. Columns A and B of that sheet will always be the list of names and the most recent prices. The main sheet can then reference this data, and no calculations and formatted fields will be removed or turned invalid by running the notebook. Lastly, a new excel-file is created in the root folder.

# Next Steps

- move from a jupyter-notebook to a .py file that we can run via a bash script to have a one-click way to get our price data.

