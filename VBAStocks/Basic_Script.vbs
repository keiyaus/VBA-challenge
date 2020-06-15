' Stock market analyst 

Sub StockAnalysis ()

' Loop through all worksheets 
For Each ws in worksheets

    ' Fill out the header for the stock analysis summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change ($)"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Set variable for holding the ticker
    Dim Ticker As String

    ' Set variable for holding the yearly stock price change
    Dim Yearly_Change As Double

    ' Set variable for holding the opening price in a year
    Dim Open_Price As Double
    Open_Price = Cells(i, 3).Value

    ' Set vaiable for holding the closing price in a year
    Dim Close_Price As Double
    Close_Price = Cells(i, 6).Value

    ' Set variable for holding the yearly percentage change in stock price
    Dim Percent_Change As Double

    ' Set variable for holding the total stock volumne
    Dim Total_Volume As Long
    Total_Volume = 0

    ' Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(x1Up).LastRow

    ' Keep track of the location fo each stock ticker in the summary table
    Summary_Table_Row = 2

    ' Loop through all the stocks within each worksheet
    For i = 2 To LastRow

        ' Check if we are still witin the same stock ticker, if not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the ticker
            Ticker = Cells(i, 1).Value
            
            ' Print the ticker in the summary table
            Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Add to the total volume
            Total_Volume = Total_Volume + Cells(i, 7).Value

            ' Print the total volume to the summary table
            Range("L" & Summary_Table_Row).Value = Total_Volume
      
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset the total volume
            Total_Volume = 0

        ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the total volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
        
            ' Calculate the yearly stock price change
            Yearly_Change = Cells(- Cells(i, 6).Value


        End If
    
    Next i

    ' Add AUTOFIT command
    ' ws.Columns("I:L").AutoFit

Next ws

End Sub