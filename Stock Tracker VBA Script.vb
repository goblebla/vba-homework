Sub Ticker()

'Variable to designate what row to print the unique output of each stock in the summary table
Dim TableRow As Long
TableRow = 2

'Variable to count all rows in a data set
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Variable to store the Ticker symbol
Dim TickerName As String

'Variable to store open price
Dim OpenPrice As Double
OpenPrice = 0

'Variable to store closed price
Dim ClosePrice As Double
ClosePrice = 0

'Variable for yearly price difference
Dim YearPriceChange As Double
YearlyVar = 0

'Variable for yearly price percent variance
Dim YearPriceVar As Double
YearPriceVar = 0

'Variable for total stock volume
Dim StockTotal As Double
StockTotal = 0

'Variables for Total Summary Table
Dim HighPercentTicker As String
HighPercentTicker = " "
Dim LowPercentTicker As String
LowPercentTicker = " "
Dim HighPercent As Double
HighPercent = 0
Dim LowPercent As Double
LowPercent = 0
Dim HighVolumeTicker As String
HighVolumeTikcer = " "
Dim HighVolume As Double
HighVolume = 0


'Summary Table Titles
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Format Table
Range("K:K").NumberFormat = "0.00%"
Range("L:L").NumberFormat = "$#,##0"

'Total Summary Table Titles
Range("O2") = "Highest % Increase"
Range("O3") = "Highest % Decrease"
Range("O4") = "Highest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

'Total Summary Table format
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "$#,##0"

'Setting correct price for first Ticker
OpenPrice = Cells(2, 3).Value


'Creating loop that will look through each stock Ticker
For I = 2 To LastRow

    StockTotal = StockTotal + Cells(I, 7)

    'Comparing the next cell value (Stock Ticker) to current cell value (Stock Ticker)
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
        'Set new values for variables
        TickerName = Cells(I, 1).Value
        ClosePrice = Cells(I, 6).Value
        
        'Calculate yearly price change
        YearPriceChange = ClosePrice - OpenPrice
        
        'Check division by 0
        If OpenPrice <> 0 Then
            YearPriceVar = YearPriceChange / OpenPrice
        Else
            YearPriceVar = 0
        End If
        

        'Print Value in Summary Table
        Cells(TableRow, 9).Value = TickerName
        Cells(TableRow, 10).Value = YearPriceChange
        Cells(TableRow, 11).Value = YearPriceVar
        Cells(TableRow, 12).Value = StockTotal
        
        'Format variance columns
        If YearPriceVar >= 0 Then
            Cells(TableRow, 11).Interior.ColorIndex = 4
        Else
            Cells(TableRow, 11).Interior.ColorIndex = 3
        End If
    
        'Comparing max values for second table
        If (YearPriceVar > HighPercent) Then
            HighPercent = YearPriceVar
            HighPercentTicker = TickerName
        ElseIf (YearPriceVar < LowPercent) Then
            LowPercent = YearPriceVar
            LowPercentTicker = TickerName
        End If
    
        If (StockTotal > HighVolume) Then
            HighVolume = StockTotal
            HighVolumeTicker = TickerName
        End If
        
        'Reset variable for next iteration
        YearPriceChange = 0
        YearPriceVar = 0
        StockTotal = 0
        ClosePrice = 0
        OpenPrice = Cells(I + 1, 3).Value
        TableRow = TableRow + 1

        'Printing new Total Summary
        Range("P2").Value = HighPercentTicker
        Range("P3").Value = LowPercentTicker
        Range("P4").Value = HighVolumeTicker
        Range("Q2").Value = HighPercent
        Range("Q3").Value = LowPercent
        Range("Q4").Value = HighVolume
    End If
Next I

End Sub

