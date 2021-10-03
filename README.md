# Refactored Stock Analysis by Year

## Project Overview

This project requires the collection of information for each stock ticker. Once this information is collected, a calculation is run to find the percentage of return. The data is then written to a sheet and formated based on the values. Green means the stock had a positive return and red means a negative return. The code should run efficiently run compared to the first iteration of the data analysis.

### Purpose

The purpose of this project is to understand the refactoring process and how it can benefit the analysis of larger sets of data.

## Results

A number of changes were made to the original code to greatly improve performance. The first optimation made was to move the data writing outside of the For loop. Here is a snip of the original code:

```
For i = 0 to UBound(tickers)
    For j = rowStart To rowEnd
        'Iterate over rows
        'Write data to sheet
        Worksheets("All Stocks Analysis").Cells(4 + i, 1).Value = ticker
        Worksheets("All Stocks Analysis").Cells(4 + i, 2).Value = totalVol
        Worksheets("All Stocks Analysis").Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    Next j
Next i
```
This was changed to:

```
For i = 0 to UBound(tickers)
    For j = rowStart To rowEnd
    'Iterate over data
    Next j
Next i
'Write data to sheet
For output = 0 To UBound(tickers)
    Worksheets("All Stocks Analysis").Cells(4 + output, 1).Value = tickers(output)
    Worksheets("All Stocks Analysis").Cells(4 + output, 2).Value = tickerVolumes(output)
    Worksheets("All Stocks Analysis").Cells(4 + output, 3).Value = (tickerEnd(output) / tickerStart(output)) - 1
Next output
```


