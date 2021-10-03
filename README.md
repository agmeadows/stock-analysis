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
For tickerIndex = 0 to UBound(tickers)
    For rowNum = rowStart To rowEnd
        'Iterate over data
    Next rowNum
Next tickerIndex
'Write data to sheet
For output = 0 To UBound(tickers)
    Worksheets("All Stocks Analysis").Cells(4 + output, 1).Value = tickers(output)
    Worksheets("All Stocks Analysis").Cells(4 + output, 2).Value = tickerVolumes(output)
    Worksheets("All Stocks Analysis").Cells(4 + output, 3).Value = (tickerEnd(output) / tickerStart(output)) - 1
Next output
```
This ensures the data is only written once for each ticker and does not slow down the iteration over the data.

The next change was to modify how the start and end prices are set. The original code was as follows:

```
For i = 0 To UBound(tickers)
        ticker = tickers(i)
        For j = rowStart To rowEnd
            'Get starting price
            If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            'Get ending price
            If Cells(j + 1, 1) <> ticker And Cells(j, 1) = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
    Next j
Next i
```
This caused the If statements to execute on every iteration of the loop which added significant time to the calculation.

 The refactored code is much simpler:

 ```
 For tickerIndex = 0 To UBound(tickers)
        ticker = tickers(tickerIndex)
        tickerStart(tickerIndex) = Cells(rowStart, 6).Value
        For rowNum = rowStart To rowEnd
            'Increase total volume
            If Cells(rowNum, 1) = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(rowNum, 8).Value
            Else
                Exit For
            End If
        Next rowNum
        tickerEnd(tickerIndex) = Cells(rowNum - 1, 6).Value
        rowStart = rowNum
Next tickerIndex
```

As you can see, the start price is set just before interating over the next row of the stock ticker. The first entry is always the start price. 

The end price is set in a similar manner. It is set after the last row of the ticker is found. the last item will always be the end price

Together, these changes improved the speed of the code by around 500%. Below is the execution time of the original code compared to the refactored code:

![Original Speed](/Resources/Runtime_orig_2018.PNG)