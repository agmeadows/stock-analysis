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

#### Original
![Original Speed](/Resources/Runtime_orig_2018.PNG)

#### Refactored
![Refactored Speed](/Resources/Runtime_refactor_2018.PNG)

The new code runs in a fraction of a second compared to nearly 30 seconds for the original

## Summary

It is clear refactoring has major benefits. Through this process, the new code can run significantly faster. This allows for the code to handle larger sets of data. The biggest drawback is the additonal time involved in optimizing. This should only be done for code that will see multiple uses and not for one-off projects.

In this case, the refactored VBA script was greatly improved. Since the code can be run for multiple years, it benefits from refactoring. Each run is significantly faster. Should it be required, additional tickers and associated could be added with minimal impact to performance. The original code was already at its limit. If any additional data was added the problem would only get worse.

The biggest challenge in refactoring is determining the best place to execute a command to avoid redundant operations. Often times, this requires in-depth thinking of the process and exactly how the code will run. Debugging tools are very helpful in finding what the code is doing and when. however, this can take quite a bit of time to do. 