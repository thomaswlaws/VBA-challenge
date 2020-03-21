Attribute VB_Name = "Module11"
Sub StockOpen():
Dim ticker_location As String
Dim lastrow As Double
Dim Summary_table_Row As Double
Dim TotalVolume As LongLong
TotalVolume = 0
Dim Year_open As Double
Dim Year_close As Double
For Each ws In Worksheets
    ws.Activate

'Set Heading Locations
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Volume Traded"
    'Get the last row noted
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Get the location for where the ticker_location info will end up
    Summary_table_Row = 2
        For i = 2 To lastrow

'Note location for yearly opening value
        Year_open = Cells(2, 3)
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                If Year_open <> 0 Then

'Note location of yearly closing value
                    Year_close = ws.Cells(i, 6)

'Calculate the Yearly Change
                    YearlyChange = Year_close - Year_open

'Calculate the Percent Change
                    PercChange = (Year_close - Year_open) / (Year_open)

'Note the location for the rest of the Important Values
                    Range("J" & Summary_table_Row) = YearlyChange
                    Range("K" & Summary_table_Row) = PercChange
                    Range("K" & Summary_table_Row).Style = "Percent"

'Get Conditional Formatting Set up
                        If Range("J" & Summary_table_Row) > 0 Then
                            Range("J" & Summary_table_Row).Interior.ColorIndex = 4
                            Range("K" & Summary_table_Row).Interior.ColorIndex = 4
                        Else
                            Range("J" & Summary_table_Row).Interior.ColorIndex = 3
                            Range("K" & Summary_table_Row).Interior.ColorIndex = 3
                        End If
                End If

'Redefining the varable for the yearly opening value
                Year_open = ws.Cells(i + 1, 6)
                ticker_location = Cells(i, 1).Value
                TotalVolume = TotalVolume + Cells(i, 7).Value
                Range("I" & Summary_table_Row) = ticker_location
                Range("L" & Summary_table_Row) = TotalVolume

'Adding 1 to the the Summary Table Row
                Summary_table_Row = Summary_table_Row + 1

'Getting the volume reset for the next
                TotalVolume = 0
            Else
                TotalVolume = TotalVolume + Cells(i, 7).Value
            End If
        Next i
Next ws
End Sub
