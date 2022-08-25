# VBA_Challenge

##Overview of Project

The VBA Challenge was a Excel project which used a existing data file containing two years of stock reports.   Creating a macro 
and using refactoring, the run times and the correct information displayed in the charge was the final objective.   The final 
chart shows the Ticker, the Total Daily volume, and the Return in percent of the stocks.   The dates for the daily volume were 
taken from January 01 till December 31 for both the 2017 stock year and the 2018 stock year.   

##Analysis and Challenge

In order to create a macro for the stocks, the correct tickers needed to be initialized.   A variable was created 
and set as a string.   Then the tickers array was created setting each ticker value.

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
    
    

