Attribute VB_Name = "Module1"
Sub StockVolume()

'Define variable for the ticker symbol
    Dim Ticker As String
    
'Define variable for the total stock volume
    Dim Stock_Volume_Total As Double

'Total stock volume starts at 0
    Stock_Volume = 0
    
'Define variable for each row in the summary table (added ticker volumes)
    Dim Summary_Table_Row As Double
    
'Summary Table begins at row 2
    Summary_Table_Row = 2
    
'Place ticker label in a cell
    Cells(1, 10).Value = "Ticker"

'Place total stock volume in a cell
    Cells(1, 11).Value = "Total Stock Volume"

'Determine the last row of the data
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Start at row 2 and loop through the data until it hits the last row
    For i = 2 To lastrow

'If the cells ticker does not equal the next ticker then do the following
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
'Add the stock volumes in column 7 (G)
            Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

'Place the ticker symbol in column J
            Range("J" & Summary_Table_Row).Value = Ticker

'Place the total stock volume in column K
            Range("K" & Summary_Table_Row).Value = Stock_Volume_Total
            
'Move to the next row in the summary table
            Summary_Table_Row = Summary_Table_Row + 1

'Start back at 0 and add up the next stock's volume
            Stock_Volume_Total = 0

        Else
            Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

        End If

    Next i

 End Sub

