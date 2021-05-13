Sub DQAnalysis()

Worksheets("DQ Analysis").Activate
Range("A1").Value = "DAQO (Ticker: DQ)"

'Create a header row
Cells(3, 1) = "Year"
Cells(3, 2) = "Total Daily Volume"
Cells(3, 3) = "Return"

Worksheets("2018").Activate

'starting from second rowto avoid header row
rowStart = 2

'DELETE: rowEnd = 3013

'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
totalVolume = 0

Dim startingPrice As Double
Dim endingPrice As Double

 
For i = rowStart To rowEnd

    'check if its the first occurance of DQ
   If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
   'find the starting price
   startingPrice = Cells(i, 6).Value

   End If
   
    'check if its the last occurance of DQ
   If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
   
   'find the ending price
   endingPrice = Cells(i, 6).Value

   End If
   
    'increase totalVolume if ticker is "DQ"
    If Cells(i, 1).Value = "DQ" Then
    
        totalVolume = totalVolume + Cells(i, 8).Value
        
    End If

Next i

'Our return values

Worksheets("DQ Analysis").Activate
Cells(4, 1).Value = 2018
Cells(4, 2).Value = totalVolume
Cells(4, 3).Value = endingPrice / startingPrice - 1

End Sub

Sub AllStocksAnalysis()

Worksheets("All Stocks Analysis").Activate
Range("A1").Value = "All Stocks (2018)"

'Create a header row
Cells(3, 1) = "Ticker"
Cells(3, 2) = "Total Daily Volume"
Cells(3, 3) = "Return"

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

Dim startingPrice As Double
Dim endingPrice As Double

Worksheets("2018").Activate
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop throught tickers
    For i = 0 To 11
    ticker = tickers(i)
        totalVolume = 0
        
        Worksheets("2018").Activate
        'loop through rows
        For j = 2 To RowCount
        
         'get the totalVolume of the ticker
          If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        
          End If
          
           'check if its the first occurance of the ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
             
             'find the starting price
            startingPrice = Cells(j, 6).Value
            End If
   
            'check if its the last occurance of the ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
         
            'find the ending price
             endingPrice = Cells(j, 6).Value

            End If
          
          Next j
          
         'Our return values
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
End Sub
