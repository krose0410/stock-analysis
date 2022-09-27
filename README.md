# **stock-analysis**
Click here to view the Excel file: VBA Challenge - Stock Analysis]( https://github.com/krose0410/stock-analysis/blob/main/stock%20analysis/VBA_Challenge.xlsm![image](https://user-images.githubusercontent.com/109715441/192431378-0082705d-23c3-42c3-b41b-1e546af7d8a1.png)

## **Overview of Project**
The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.
### **Results**
#### The Data
The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock. 

<img width="278" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/109715441/191660673-715ca9d4-c068-4836-8414-ddd65ed6cabc.png">
<img width="273" alt="Screen Shot 2022-09-21 at 8 57 17 PM" src="https://user-images.githubusercontent.com/109715441/191660692-643a4368-1e67-448d-8945-677957b27900.png">

##### The Anaylysis
Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out to set the structure for the refactoring. Below is the instruction and code as written in the file.  '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
## **Summary**
###### Pros and Cons of Refactoring
Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

###### Advantages of Refactoring
The biggest benefit that occurred because of the refactoring was a decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. Attached below are the screenshots that indicate the run time for our new analysis.
