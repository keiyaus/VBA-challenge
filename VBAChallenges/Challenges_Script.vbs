Sub StockChallenges()

    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Fill out header for new table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Determine the last row in stock analysis summary table
        Table_LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row

        ' Set the needed variables
        Dim GrtIncrease As Double
        Dim GrtIncrease_Row As Double

        Dim GrtDecrease As Double
        Dim GrtDecrease_Row As Double

        Dim GrtVolume As Double
        Dim GrtVolume_Row As Double

        ' Find and print the values to new table
        GrtIncrease = WorksheetFunction.Max(ws.Range("K1:K" & Table_LastRow))
        ws.Range("Q2").Value = GrtIncrease

        GrtDecrease = WorksheetFunction.Min(ws.Range("K1:K" & Table_LastRow))
        ws.Range("Q3").Value = GrtDecrease

        GrtVolume = WorksheetFunction.Max(ws.Range("L1:L" & Table_LastRow))
        ws.Range("Q4").Value = GrtVolume

        For i = 2 To Table_LastRow

            If ws.Cells(i, 11).Value = GrtIncrease Then
                ws.Range("P2").Value = ws.Cells(i, 9)

            ElseIf ws.Cells(i, 11).Value = GrtDecrease Then
                ws.Range("P3").Value = ws.Cells(i, 9)

            End If

            If ws.Cells(i, 12).Value = GrtVolume Then
                ws.Range("P4").Value = ws.Cells(i, 9)
            End If
            
        Next i

    ' Format the new columns
    ws.Columns("O:Q").AutoFit
    
    Next ws


End Sub