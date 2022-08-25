# VBA_Challenge

##Overview of Project

The VBA Challenge was a Excel project which used a existing data file containing two years of stock reports.   Creating a macro 
and using refactoring, the run times and the correct information displayed in the charge was the final objective.   The final 
chart shows the Ticker, the Total Daily volume, and the Return in percent of the stocks.   The dates for the daily volume were 
taken from January 01 till December 31 for both the 2017 stock year and the 2018 stock year.   

##Analysis and Challenge

In order to create a macro for the stocks, the correct tickers needed to be initialized.   A variable was created 
and set as a string.   Then the tickers array was created setting each ticker value.

 
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
    
    
    The worksheet(yearValue) was made active and the Row count was determined by using the boiler plate below.
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row 
    
    
   The tickerIndex variable was created and then initialized as zero.  Three other variables tickerVolumes,
   tickerStartPrices, and tickerEndingPrices were created and set with the correct data type.
   
    Dim tickerIndex As Integer
    
    tickerIndex = 0
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
   A for loop was then created to initialize the tickerVolumes to run through the ticker names on the spreedsheet.
   
     For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    Next i
    
   In order to loop through all of the rows in the spreedsheet, a for loop was created.  As the rows were loped through, 
   tickerVolumes was added to itself to increase the total tickerVolume of the last column named Volume.
   
    For i = 2 To RowCount
    
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
   Two If statements were used to find the start and end of each tickerIndex.  The first If found the start of the ticker in
   in the spreadsheet and the second If found the end.  Once the end of the tickers was found, the tickerIndex was added 
   to move to the next ticker.
   
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
       
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
    End If
    
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        tickerIndex = tickerIndex + 1
        
   The active worksheet was changed to All Stocks Analysis.   The chart for the macro was created using the assigned cells 
   for the headers.
   
   
    Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = ((tickerEndingPrices(i) - tickerStartingPrices(i)) / tickerStartingPrices(i))
    Next i
    
    
   The chart cells value were set using the Range and format below.
   
    
    Worksheets("All Stocks Analysis").Activate
     Range("A3:C3").Font.FontStyle = "Bold"
     Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
     Range("B4:B15").NumberFormat = "#,##0"
     Range("C4:C15").NumberFormat = "0.0%"
     Columns("B").AutoFit   
     
   
   The colors of the cell types were created using a for loop.   
   

     dataRowStart = 4
     dataRowEnd = 15
   
     For i = dataRowStart To dataRowEnd
     
        If Cells(i, 3) > 0 Then 
           Cells(i, 3).Interior.Color = vbGreen    
        Else
           Cells(i, 3).Interior.Color = vbRed
        End If    
     Next i
     
    
    
    In order to calculate the elapsed time for running the macro, the following code was used.   Here we have the dataRowsStart
    and dataRowsEnd.
   
   
   
 
      endTime = Timer
      MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    
    

