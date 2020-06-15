'Problem:'
'Create a loop that scans through a years worth of stocks'
'***** MODIFY? THIS SCRIPT TO GO THROUGH ALL YEARS (SHEETS) *****


Sub StockTickerScanCompare()

' Set an initial variable for holding the stock name
  Dim Stock_Name As String

  ' Set an initial variable for holding the volume of a stock over years'
  Dim Yr_Yr_stock_vol_Total As Double
  Yr_Yr_stock_vol_Total = 0

'Keep track of the location for each stock name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


    'scan through stock column A'
    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all stocks
    For i = 2 To 70926

        ' Check if we are still within the same stock, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the Stock name
        Stock_Name = Cells(i, 1).Value

        ' Add to the Stock Total
        Stock_Total = Stock_Total + Cells(i, 3).Value

        ' Add the name Ticker to the Column Header for "I"
        Cells(1, 9).Value = "Ticker"

        ' Print the Stock Name in the Summary Table
        Range("I" & Summary_Table_Row).Value = Stock_Name

        ' Print the Stock Name amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = Yr_Yr_stock_vol_Total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Brand Total
        Yr_Yr_stock_vol_Total = 0

        ' If the cell immediately following a row is the same stock...
    Else

        ' Add to the stock Total
        Yr_Yr_stock_vol_Total = Yr_Yr_stock_vol_Total + Cells(i, 3).Value

    End If

  Next i

  '** Color change values based on Positvive or Negative Growth**

  'Sample color change scripts. *** ADD conditionals based on positve of negative
  'change'

  ' Set the Font color to Red
  'Range("A1").Font.ColorIndex = 3

  ' Set the Cell Colors to Red
  'Range("A2:A5").Interior.ColorIndex = 3

  ' Set the Font Color to Green
  'Range("B1").Font.ColorIndex = 4

  ' Set the Cell Colors to Green
  'Range("B2:B5").Interior.ColorIndex = 4

End Sub
