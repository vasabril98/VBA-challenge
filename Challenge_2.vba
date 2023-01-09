Sub stock_parser()

    For Each ws in Worksheets

        ' Setup all your variables
        Dim LastRow, LastCol As Long
        Dim Summary_Table_Row As Integer
        Dim Open_Price_Year, Close_Price_Year, Total_Stock_Volume, Percent_Change, Greatest_Increase, Greatest_Decrease, Greatest_Total_Volume  As Double
        Dim Ticker_Symbol, Greatest_Increase_Ticker, Greatest_Decrease_Ticker As String
        Dim WorksheetName As String

        ' Find the last rows and columns of th table
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Assign all your variables to the corresponding values
        Summary_Table_Row = 2
        Total_Stock_Volume = 0
        ws.Cells(1, LastCol + 2).Value = "Ticker"
        ws.Cells(1, LastCol + 3).Value = "Yearly Change"
        ws.Cells(1, LastCol + 4).Value = "Percent Change"
        ws.Cells(1, LastCol + 5).Value = "Total Stock Volume"
        ws.Cells(2, LastCol + 8).Value = "Greatest % Increase"
        ws.Cells(3, LastCol + 8).Value = "Greatest % Decrease"
        ws.Cells(4, LastCol + 8).Value = "Greatest Total Volume"
        ws.Cells(1, LastCol + 9).Value = "Ticker"
        ws.Cells(1, LastCol + 10).Value = "Value"

        Open_Price_Year = ws.Range("C2").Value
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Total_Volume = 0
        Greatest_Increase_Ticker = ""
        Greatest_Decrease_Ticker = ""
        
        ' Loop that iterates over all the rows of the table
        For i = 2 To LastRow
            ' Conditional that checks if the ticker symbol of the next row is different
            ' if it is then performs the following operations
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Grabs the needed values from the current row
                Close_Price_Year = ws.Cells(i, 6).Value
                Ticker_Symbol = ws.Cells(i, 1).Value
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i,7)
                Percent_Change = (Close_Price_Year - Open_Price_Year)/Open_Price_Year

                ' Calculates and assign the values to each column
                ws.Cells(Summary_Table_Row, LastCol + 2).Value = Ticker_Symbol
                ws.Cells(Summary_Table_Row, LastCol + 3).Value = Round(Close_Price_Year - Open_Price_Year,2)
                ws.Cells(Summary_Table_Row, LastCol + 4).Value = FormatPercent(Percent_Change, 2)
                ws.Cells(Summary_Table_Row, LastCol + 5).Value = Total_Stock_Volume
                
                ' Conditional that assigns the corresponding cell color (red = 3, green = 4) to the "Yearly Change" column 
                If ws.Cells(Summary_Table_Row, LastCol + 3).Value > 0 Then
                    ws.Cells(Summary_Table_Row, LastCol + 3).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, LastCol + 3).Interior.ColorIndex = 3
                End If

                ' Conditional that saves the current Percent_Change value if it is the largest or the smallest
                If Greatest_Increase < Percent_Change Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = Ticker_Symbol
                ElseIf Greatest_Decrease > Percent_Change Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = Ticker_Symbol
                End If

                ' Conditional that saves the Total_Stock_Volume value if it is the largest
                If Greatest_Total_Volume < Total_Stock_Volume Then
                    Greatest_Total_Volume = Total_Stock_Volume
                    Greatest_Total_Volume_Ticker = Ticker_Symbol
                End If

                ' Resets the values for the calculations of the next Stock
                Open_Price_Year = ws.Cells(i + 1, 3).Value
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                
            ' If the ticker symbol of the next row is the same then performs the following operations
            Else
                ' Add to the current volume to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i,7)
            End If
        Next i

        ws.Cells(2, LastCol + 9).Value = Greatest_Increase_Ticker
        ws.Cells(3, LastCol + 9).Value = Greatest_Decrease_Ticker
        ws.Cells(4, LastCol + 9).Value = Greatest_Total_Volume_Ticker
        
        ws.Cells(2, LastCol + 10).Value = FormatPercent(Greatest_Increase,2)
        ws.Cells(3, LastCol + 10).Value = FormatPercent(Greatest_Decrease,2 )
        ws.Cells(4, LastCol + 10).Value = Greatest_Total_Volume

    Next ws

End Sub

