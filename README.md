# Stock_Analysis

## Overview of the project

Steve found the Daily volume and yearly return for each each stock. Although the code works well for a dozen stocks it might not work as well for thousands of stocks.
And if it does, it may take along time to execute. In this project we are editing or refactoring the code to loop through all the data one time inorder to collect the same information. 


## Results
When comparing the stocks 2017 and 2018. Overall performance of 2017 is better than 2018. The returns are formatted to show green and red color. So it will be easy to identify which year and stocks performed well.
      The code refactured and gives the same result. We had a timer set in the code to find the code running time. And its found for 2017 after refacturing the code, the code ran 4 times faster than previous code and for 2018 code ran 4.6 times faster than the previous code. This a great proof that a code can be written in multiple ways and its important to write the code in most efficicent way.


   _code run time for 2017_
   
  
   ![2017](https://github.com/11nithin/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png.PNG) 
   --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   _code run time for 2017 and 2018_
   
   ![2018](https://github.com/11nithin/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png.PNG)
   
  
  _Refactored code_
  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
````
   '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
    tickerVolumes(j) = 0
    Next j
            
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                 
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
                       
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
       
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
             End If
             
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
                         
        'End If
            End If  
             Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
     Cells(4 + i, 1).Value = tickers(i)
     Cells(4 + i, 2).Value = tickerVolumes(i)
     Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
     
        Next i
  ````      
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------       
## Summary
### Advantage of Refactoring the code in general 
 - Code runs  faster and efficiently 
 - Modifying the code in future is easy
 - All users can easily understand the code and troubleshoot if ran into any issue

### Disadvantage of Refactoring the code in general
 - Can be Time consuming 
 - Can run into error
### Advantage and Disadvantage of Refactoring the Orginial Code
- Our target was to run the code efficiently for thousands of stocks. Refactoring the original code made the code run 4 times faster. For larger dataset it would be very useful. 
- At the same time it was time consuming and ran into several errors while refactoring. But it works better for large amount of data.
