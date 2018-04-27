Sub Stocks()

    Set ws = ThisWorkbook.ActiveSheet
    
    For Each ws In Worksheets
        
        'Fill in the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Find the last row
        Dim Last_Row As Long
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Keep a counter for different stocks
        Dim Count As Long
        Count = 0

        'keep a counter for the first of each stock
        Dim First_of_each As Long

        'keep a counter for the last of each stock
        Dim Last_of_each As Long

        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        Dim Open_Value As Double

        For i = 2 To Last_Row

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                Count = Count + 1

                'Fill in the different tickers in each row
                Dim Ticker As String
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & (Count + 1)).Value = Ticker

                'Add the Stock Volume of each row to Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                'Fill in the Total Stock Volume in each row
                ws.Range("L" & (Count + 1)).Value = Total_Stock_Volume

                'Reset the Total_Stock_Volume to 0 for the next stock
                Total_Stock_Volume = 0

                'Store the Close Value
                Dim Close_Value As Double
                Close_Value = ws.Cells(i, 6).Value

                'Fill in the Yearly Change
                ws.Range("J" & (Count + 1)).Value = Close_Value - Open_Value

                'Fill in the Percent Change
                Dim Percent_Change As Double
                If Open_Value <> 0 Then
                    Percent_Change = (Close_Value - Open_Value) / Open_Value
                    ws.Range("K" & (Count + 1)).Value = Percent_Change
                End If
                ws.Range("K" & (Count + 1)).Style = "Percent"

            Else
                'Add the Stock Volume of each row to Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                'Check if the row is a diffrent stock from the rows above it
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                    'Store the open value for this row
                    Open_Value = ws.Cells(i,
                     3).Value

                End If

            End If

        Next i
               
        For n = 2 To Count + 1
            If ws.Range("J" & n).Value < 0 Then
                ws.Range("J" & n).Interior.ColorIndex = 3
            ElseIf ws.Range("J" & n).Value > 0 Then
                ws.Range("J" & n).Interior.ColorIndex = 4
            End If
        Next n

    Next ws
    
End Sub




















