# VBA Challenge

## Overview of Project

For this project, an edited, or refactored, code that was created to analyze stocks from 2017 and 2018 was updated so that it can be used on stock data over the last few years. In the refactored code, no new functionality is added but the code is made more efficient so that it can be used on a larger data sets.

## Results

Since the purpose of this project was to use code that was previously written, I copied the code that was used to create the header, ticker array, and input box. After that, I refactored my code to make it more efficient making the code applicable to larger data. Below is an example of the original code and refactored code:

Original Code:

5a.Find the total volume for the current ticker.

If Cells(j, 1).Value = ticker Then
   totalVolume = totalVolume + Cells(j, 8).Value
End If

5b.Find the starting price for the current ticker.

If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
  startingPrice = Cells(j, 6).Value
End IF

5c.Find the ending price for the current ticker.

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
  endingPrice = Cells(j, 6).Value
  End If
  
Refactored Code
''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
       
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i


## Summary

The advantages of refactoring code are making the code more efficient by removing redundant code, using less memory and improving the logic of the code to make it easier for users to read. A disadvantage of refactored codes is making errors that can lead to errors.

#### Refactored Code

Based on the times it took this data to run, the refactored code took less time to run. Below are screenshots of the time it took the original code for 2017 and the refactored code for 2017.

![Screen Shot 2022-06-16 at 3 14 32 PM](https://user-images.githubusercontent.com/85198012/174149010-c7b6b3f8-2928-47c6-9eef-d8923fa3bb08.png)
 
 Screenshot of Time for original Code for 2017
 
![Screen Shot 2022-06-16 at 3 08 52 PM](https://user-images.githubusercontent.com/85198012/174149113-8eb8299a-b498-42d0-9f52-08430459603b.png)

Screenshot of Time for Refactored Code for 2017

