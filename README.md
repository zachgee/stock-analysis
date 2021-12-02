# Green Stock Analysis

### Overview

For this project we were asked to help a dear friend, who himself was asked by his parents to help understand how their investment in the Green energy company DQ is doing. To help our friend, we looked at how different green company stocks performed in the year 2017 and compared those results to how the stocks performed in 2018. Our friend provided us with the raw data through a Microsoft Excel sheet - and we used our VBA skills to perform the analysis. Once a code was created, we also decided to refractor the code in order to prove how cleaning up code can make a code run faster and more efficiently. 

### Results

The results we found may not be the best news for our friends parents - but it is best they know how their stock are performing as opposed to being left in the dark. The following VBA code was used in order to find our results: 

"Sub AllStocksAnalysis()
 Dim startTime As Single
 Dim endTime  As Single

yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

'1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   'color the cells
   
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    
End Sub"

The following images show the reult of our code:

![Results_2017](https://user-images.githubusercontent.com/87949792/144359929-17bed1c9-6850-4fee-8f0b-48f8ce155ed8.png)

![Results_2018](https://user-images.githubusercontent.com/87949792/144360003-3a5f8a97-9dcd-4904-8540-17d9c49e5508.png)

Lastly, in order to refractor our code, the following VBA script was used: 

"Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Title Analysis
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
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    '2a) Create for loop to analyze each ticker in the array.
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End If
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End if
            End If
            
        '3d) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                'if it is, increase tickerIndex to move on to next ticker in array.
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'ReturnValue
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

    
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub"
The following screenshots show the run time of the original code:

![Timer_Original_Code_2017](https://user-images.githubusercontent.com/87949792/144360387-ecc19b06-fc03-4d98-ad2f-f47195f6699a.png)


![Timer_Original_2018](https://user-images.githubusercontent.com/87949792/144360421-e0f2cf57-4e58-42ed-94f5-42214a10737c.png)


The following screenhots show the refractored codes' run times:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/87949792/144360208-ce310e99-3476-4fc8-8a4b-252f6d6edaff.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/87949792/144360219-121d7bd0-4baf-4af4-bf33-72b027d48c66.png)




### Analysis

Unfortunately for our friends parents, the majority of the company stocks anylized saw a decrease in ruturn, their DQ stocks included.

From a coding perspective however, we have made a good case for refractoring code. As the images show, the run times dropped a decent ammount after being refractored. However, this does not come without some cons - refractoring code can get messy and created room for unwanted errors. 


