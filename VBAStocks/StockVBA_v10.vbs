'Problem:'
'Create a loop that scans through a years worth of stocks'
'***** MODIFY? THIS SCRIPT TO GO THROUGH ALL YEARS (SHEETS) *****


Sub StockTickerScanCompareREV7()

'****NOTE removed ws as...... for this script******
    'set current Worksheet as ws 
     'Dim ws As Worksheet
     
     
     '***USE IF ' Loop through all of the worksheets in the active workbook. **cite Ibaloyan**
    
        'For Each ws In Worksheets
  
  'Set an initial variable for holding the stock name
  Dim Stock_Name As String
  Stock_Name = ""

  ' Set an initial variable for holding the volume of a stock over years'
  Dim Yr_Yr_stock_vol_Total As Double
  Yr_Yr_stock_vol_Total = 0

' Keep track of the location for each ticker name
' in the summary table for the current worksheet
 'Keep track of the location for each stock name in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  'Dim CompStock As Range  TESTING*****
'set variables for price changes" **cite Ibaloyan**
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0
        
     

 'Create an Array to hold the tickers found in a loop pass, to find the First/Last entry'**TESTING
 'Dim TickerFirstLastArray() As Single
 'Dim TickerFirstRow As Integer
 'Dim TickerLastRow As Integer
 


'Change COLUMN HEADER SECTION'
        ' Add the name Ticker to the Column Header for "I"
        Cells(1, 9).Value = "Ticker"
        ' Add the name Year Change to the Column Header for "J"
        Cells(1, 10).Value = "Year Change"
        ' Add the name Percent Change to the Column Header for "K"
        Cells(1, 11).Value = "Percent Change"
        ' Add the name Total Stock Volume to the Column Header for "L"
        Cells(1, 12).Value = "Total Stock Volume"

    'AutoFit G-L Columns on Worksheet, I thought this may be a nice cleanup'
    Range("I1:L1").Columns.AutoFit

    'Find unique Stock Ticker Year Price Change' TESTING*****
     'Set CompStock = Range(“A:A”).Find (What:=Range(“A2”).Value, LookIn:=xlValues, LookAt:= xlWhole)
     'MsgBox CompStock

    'scan through stock column A' TESTING******
    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Set initial value of Open Price for the first Ticker of CurrentWs,
        ' The rest ticker's open price will be initialized within the for loop below
        Open_Price = Cells(2, 3).Value

    ' Set initial row count for the current worksheet
    Dim Lastrow As Long
    Dim i As Long
        
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all stocks
    For i = 2 To Lastrow

        ' Check if we are still within the same stock, if it is not...
        'If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Stock name
             Stock_Name = Cells(i, 1).Value

            ' Calculate Delta_Price and Delta_Percent
            Close_Price = Cells(i, 6).Value
            Delta_Price = Close_Price - Open_Price
            ' Check Division by 0 condition
            If Open_Price <> 0 Then
            Delta_Percent = (Delta_Price / Open_Price) * 100
            Else
            ' Unlikely, but it needs to be checked to avoid program crushing
            MsgBox ("For " & Stock_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
            End If

            'Add to the Stock Volume Total
             Yr_Yr_stock_vol_Total = Yr_Yr_stock_vol_Total + Cells(i, 7).Value

        ' Print the Stock Name in the Summary Table
        Range("I" & Summary_Table_Row).Value = Stock_Name


          ' Print the Ticker Name in the Summary Table, Column I **cite Ibaloyan**
                Range("J" & Summary_Table_Row).Value = Delta_Price
                ' Fill "Yearly Change", i.e. Delta_Price with Green and Red colors
                If (Delta_Price > 0) Then
                    'Fill column with GREEN color - good
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                    'Fill column with RED color - bad
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")

        
        
        ' Print the Stock Name amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = Yr_Yr_stock_vol_Total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Stock Volume Total
        Yr_Yr_stock_vol_Total = 0

        ' If the cell immediately following a row is the same stock...
    Else

        ' Add to the stock Total
        Yr_Yr_stock_vol_Total = Yr_Yr_stock_vol_Total + Cells(i, 7).Value

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
