# VBA Challenge

## Overview of Project

### Purpose

#### Refactor the VBA code previously used for stock parsing. This, to evaluate the refactoring and see if it would make correcting the code a faster running process. To finally evaluate again 12 different actions in a period of 2 years and establish the changes that were made in the code and how this affected the process.

## Results

### Compare stock performance between 2017 and 2018

#### When comparing the performance of the 12 stocks for both periods (2017 and 2018), it is more than evident that 2017 was a very good year compared to 2018, since only the stock called TERP has a negative performance (even not so alarming) of 7.2%. Meanwhile, for 2018 all actions have suffered a brutal drop in performance, with the exception of the action called RUN, which grew from 5.5% to 84%. It is only this stock and ENPH that are still in the green for 2018, however, the latter also suffered a huge drop from 129.5% in 2017 to 81.9% in 2018. The rest of the stocks are in red numbers with figures that amount to 60% negative performance.

##### To demonstrate this graphically, there are images of the performance of these actions for both years.

![2017](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/images/2_2017.png)

![2018](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/images/2_2018.png)

##### The code used to get these stock performance numbers is as follows:

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        tickerindex = 0
    For i = 0 To 11
        tickerindex = tickers(i)
    
        '1b) Create three output arrays
        
        Dim tickerVolumes As Long
        Dim tickerStartingPrices As Single
        Dim tickerEndingPrices As Single
        
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = tickerindex Then
        
            tickerVolumes = 0
    
            End If
            
         Next j
            
        ''2b) Loop over all the rows in the spreadsheet.
        
        For k = 2 To RowCount
        
            '3a) Increase volume for current ticker
            
            If Cells(k, 1).Value = tickerindex Then
            
                tickerVolumes = tickerVolumes + Cells(k, 8).Value
                
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
           
            If Cells(k - 1, 1).Value <> tickerindex And Cells(k, 1).Value = tickerindex Then
            
                tickerStartingPrices = Cells(k, 6).Value
                
            End If
            
            '3c) check if the current row is the last row with the selected ticker.
            
             If Cells(k + 1, 1).Value <> tickerindex And Cells(k, 1).Value = tickerindex Then
             
                tickerEndingPrices = Cells(k, 6).Value
                
             End If
             
            '3d) Increase the tickerIndex.
             
            If Cells(k + 1, 1).Value <> tickerindex And Cells(k, 1).Value = tickerindex Then
            
                'tickerindex = tickerindex + 1
                
            End If
           
        Next k
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickerindex
        Cells(4 + i, 2).Value = tickerVolumes
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


### Compare the execution times of the original script and the refactored script

#### When we compare the execution times of the old code and the refactored code, it is evident that for both years there is an increase of approximately 0.10 seconds for each case. It can be evidenced in the following images the difference that exists between the running times of both codes. These differences may not be obvious or substantial to us as humans, because we may not realize it because of how fast both codes run. However, for the machine in the case where it has to run millions of actions and maybe several years, the differences between the two codes may begin to differ. Despite this, in our case the refactored code seems to have a delay compared to the original code, but in the case of having many more actions, the same may not happen.

#####Refactored

![2.2017](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

#####Original

![1.2017](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/images/1VBA_Challenge_2017.png)

#####Refactored

![2.2018](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

#####Original

![1.2018](https://github.com/jjgomezg17/stock-analysis/blob/main/resources/images/2VBA_Challenge_2018.png)

## Summary

### What are the advantages or disadvantages of refactoring code?

#### In general, some of the advantages that would come from refactoring code would be to make it more optimal for the function or purpose for which it was created, to make it more maintainable and reusable over time, to make it more agile when running to use less machine memory. and to be able to execute functions with large databases in short times. Also, by refactoring the code it could be made more encapsulated and therefore easy to reuse in other cases and it could be that by encapsulating the code it could have fewer bugs and even vulnerabilities.

####Taking into account the above, we should also take into account some disadvantages of refactoring code, which would be the same situations as the advantages, new bugs could be created in the code when trying to make it shorter, it could also make the code more confusing by trying to make it more compact and skipping steps or making them more condensed and you could create more complex and difficult to replicate structures.

### How do these pros and cons apply to refactoring the original VBA script?

#### In the particular case of our code, trying to refactor some advantages and disadvantages were found. Among the advantages of refactoring, it was found that it is easier to execute from the VBA function a compact code that is within the same Sub, compared to three different ones in the original code. Also, it is possible that by only having to execute the same Sub in each year, if you have a database with millions of actions, it may run faster than the original, not as in the current case of only 12 actions. Finally, it was found as an advantage to condense the code and perhaps make it more logical when following the IF line, limiting the lines of code and making it shorter.

#### Taking into account the advantages found, different disadvantages were also found. Among them, the most obvious would be the fact of increasing the run time when refactoring the code, something that perhaps would not have been expected when condensing the code but that was evident when executing it. This drawback may be eliminated if you have a large number of actions but cannot check against the current database. The other disadvantage found when refactoring the code was the simple fact of having to rethink it and go through all the bugs found when redoing it, when there was already a functional code.
