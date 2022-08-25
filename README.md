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
   
    
    
    
    

