# VBA of Wallstreet

## Overview of Project

### Steve, a newly graduated financial analyst, was asked by his first client to advise them on the best green energy stock in which to invest. Providing the data for 12 green energy stocks, Steve asked for a user-friendly analysis tool to run in excel for any year he chooses. Visual Basic for Applications was chosen to code this analytical tool for Steve.

## Results

### Stock Performance

#### In 2017, the stock analysis shows that all stocks except TERP had a positive rate of return. The highest performing stocks in this year were DQ and SEDG, respectively (refer to Table 1). Contrarily, in 2018, all but two stocks had a negative rate of return with DQ performing the worst (refer to Table 2). The only stocks with a positive rate of return were RUN and ENPH, respectively (refer to Table 2). Unfortunately, looking simply at the highest rate of return would not provide Steve’s clients with an easy answer for their investments. 
Table 1
 
![image](https://user-images.githubusercontent.com/102757676/163580576-e9e78cfd-0e14-4dd8-a7a4-8acd7d621f9a.png)

Table 2
 
![image](https://user-images.githubusercontent.com/102757676/163580593-d11c0e5b-b78e-466b-9ec5-44b21fd11704.png)


### Code Performance

#### While the original code seems to run rather quickly (i.e., approximately 0.6 seconds), the refactored code runs nearly a factor of 100 times more quickly (i.e., approximately 0.006 seconds). In the original code, there is a nested for loop which loops through the 3013 rows of data 12 times, once for each ticker. Refer to the nested for loops in the below lines of code:
For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        Sheets(yearValue).Activate
        For j = 2 To RowCount
            
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            
            End If
        
        Next j

#### In contrast, the refactored code is broken into multiple, separate for loops and uses tickerIndex in arrays for tickerVolume, tickerStartingPrice, and tickerEndingPrice. This allows for the 3013 rows of data to be looped through only once and the output of each of the 12 tickers recorded appropriately. Refer to the below lines of code:
For i = 0 To 11
       
            tickerVolumes(i) = 0
        
 Next i
     
  For i = 2 To RowCount
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
                
            End If
            
  Next i


## Summary

### Refactoring in General

#### Refactoring code in general is a good practice for data analysts. Ensuring your code works on a small scale is great, but making the code quicker will allow for speedy results on larger data sets. Refactored code is often-times a more elegant and simple solution that can potentially minimize errors. Original code may have hard-coded values or nested for loops that can clog up analysis or throw errors when new data is added to the original data set. Ensuring code is useful no matter how many rows of data, or how many years over which analysis can be run is a huge benefit to the data analyst in the long run. 

### Refactoring the Code for Steve's analysis

#### Although the data set is small (only a bit over 3000 rows for 12 stocks), should Steve have a larger data set in future, the refactored code would provide Steve with results noticeably more quickly than the original code. If more stocks were added, there are only a few spots in the code that would need to be altered to accommodate these stock ticker additions. Furthermore, the refactored code allows Steve to analyze stock data from any year without making any changes to the code provided he has the data for those years. 
