Sub StockAnalysis()

    ' --------------------------------------------------
    ' Retrieve Ticker Name & Annual Trading Volume
    ' --------------------------------------------------
    
    ' Loop through all worksheets
    For Each ws In Worksheets

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
        
        ' Set vaiable for holding the closing price in a year
        Dim Close_Price As Double
        
        ' Set variable for holding the yearly percentage change in stock price
        Dim Percent_Change As Double

        ' Set variable for holding the total stock volumne
        Dim Total_Volume As Double
        Total_Volume = 0

        ' Determine the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Keep track of the location fo each stock ticker in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2

        ' Loop through all the stocks within each worksheet
         For i = 2 To last_row

            ' Check if we are still witin the same stock ticker, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the ticker
                Ticker = ws.Cells(i, 1).Value
            
                ' Print the ticker in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                ' Add to the total volume
                Total_Volume = Total_Volume + Cells(i, 7).Value

                ' Print the total volume to the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
      
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset the total volume
                Total_Volume = 0

            ' If the cell immediately following a row is the same ticker...
            Else

                ' Add to the total volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
                ' Calculate the yearly stock price change
                'Yearly_Change = Cells(- Cells(i, 6).Value

            End If
    
        Next i

        ' --------------------------------------------------
        ' Calculate Yearly Change in Price & Percentage
        ' --------------------------------------------------

        ' Reset Summary Table row to run the next loop
        Summary_Table_Row = 2

        Dim Ticker1stRow As Double
        Dim TickerLastRow As Double
    
        Dim TickerOpenPrice As Double
        Dim TickerClosePrice As Double
        
        Dim j As Long

        ' Loop through all tickers again
        For i = 2 To last_row

            ' If ticker name doesn't equal to the name in the previous row but the same as the next row, that's the starting row for each ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                Ticker1stRow = i
                    
                j = i
                    ' Loop from Ticker1stRow to find the first row that contains actual trading volume
                    While ws.Cells(j, 7).Value = 0
                        j = j + 1
                    Wend 

                    ' This is the opening price
                    TickerOpenPrice = ws.Cells(j, 3).Value
                    
            ' If ticker name equals to the name in the previous row but differs from the row after it, that's the ending row for each ticker    
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value And ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                TickerLastRow = i

                    j = i
                    
                    ' Loop from TickerLasRow to find the first row that contains actual trading volume
                    While ws.Cells(j, 7).Value = 0
                        j = j - 1
                    Wend
                    
                    ' This is the closing price
                    TickerClosePrice = ws.Cells(j, 6).Value
                
                    ws.Range("J" & Summary_Table_Row).Value = TickerClosePrice - TickerOpenPrice
                    ws.Range("K" & Summary_Table_Row).Value = FormatNumber((ws.Range("J" & Summary_Table_Row).Value / TickerOpenPrice * 100), 2) & "%"
                                           
                    Summary_Table_Row = Summary_Table_Row + 1
                
            End If
           
        Next i
       
        ' Autofit all new columns
        ws.Columns("I:L").AutoFit

        ' --------------------------------------------------
        ' Format Yearly Change in Colors
        ' --------------------------------------------------

        ' Determin the last row in Yealy Change column
        YC_LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

        ' Loop through Yearly Change column to check for negative and positive values
        For i = 2 To YC_LastRow

            ' Format positive & negative yearly changes in price with green & red respectively
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else   
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End if

        Next i

    Next ws

End Sub

