Attribute VB_Name = "StockAnalysisHard_module"
  Sub StockAnalysisHard()
  
  ' Set ws as a worksheet object variable.
    Dim ws As Worksheet
    
    'Check to see and if add Summary table headers
    Dim Check_Summary_Table_Header As Boolean

    'For the Hard Solution
    Dim LANDING_SHEET As Boolean
    
    'To set and look for a Summary Table on the current worksheet
    Check_Summary_Table_Header = False

    'We are on the loaded sheet
    LANDING_SHEET = True              'Hard Solution Check if

    
    
    'Loop through all of the worksheets available, *This possibly could've been a *DO LOOP*
    For Each ws In Worksheets

        'ws.Activate '****Testing 'Don't think I needed
    
        'Set initial variable for holding the ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        'Set an initial variable for holding the total per ticker name
        '***I wanted to Dim As Double, but I get a run-time error 6. I believe this is Mac related.
        Dim Total_Ticker_Volume As Variant
        Total_Ticker_Volume = 0
        
        'variable for year start and year end price and percentage
        
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim YearOver_Price As Double
        YearOver_Price = 0
        Dim YearOver_Percent As Double
        YearOver_Percent = 0
       
        ' Set new variables for Hard Solution
        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        'normally I would Dim this max volume as Double but being on a mac I'm going with Variant
        Dim MAX_VOLUME As Variant
        MAX_VOLUME = 0
         
        'Keep track of the location for each ticker name
        'in the summary table for the current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        
        Dim Lastrow As Long
        Dim i As Long
        'the famous, find the bottom of a column data
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'For all worksheet except the current ws
        If Check_Summary_Table_Header Then
            'Summary Table for current ws
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"

            

         Else
            'This is the first, resulting worksheet, reset flag for the rest of worksheets

            Check_Summary_Table_Header = True
         End If
        
         '***AutoFit I-L Columns on Worksheet, I thought this may be a nice cleanup'
         ws.Range("I1:L1").Columns.AutoFit
       
         'Set initial value of Open Price for the first Ticker of ws,
         'The rest ticker's open price will be initialized within the for loop below
         Open_Price = ws.Cells(2, 3).Value
        
         'Loop from the beginning of the current worksheet(Row2) till its last row
         For i = 2 To Lastrow
        
      
             'Check if we are still within the same ticker name,
             'if not - write results to summary table
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the ticker name, we are ready to insert this ticker name data
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Calculate YearOver_Price and YearOver_Percent
                Close_Price = ws.Cells(i, 6).Value
                YearOver_Price = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    YearOver_Percent = (YearOver_Price / Open_Price) * 100
                Else
                    '**FOUND This warning box ONLINE**' seems worth using if you want to avoid a spreadsheet crash
                    'MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")ееее    
                End If
                
                'Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
              
                
                'Print the Ticker Name in the Summary Table, Column I
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                'Print the Ticker Name in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = YearOver_Price
                'Fill "Yearly Change" Green and Red colors
                If (YearOver_Price > 0) Then
                    'Fill column with GREEN color for positive change
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (YearOver_Price <= 0) Then
                    'Fill column with RED color nor negative change
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 'Print the Ticker Name in the Summary Table, Column I
                ws.Range("K" & Summary_Table_Row).Value = (CStr(YearOver_Percent) & "%")
                'Print the Ticker Name in the Summary Table, Column J
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                'Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                'Reset YearOver_rice and YearOver_Percent holders, as we will be working with new Ticker
                YearOver_Price = 0
                'Hard part,do this in the beginning of the for loop YearOver_Percent = 0
                Close_Price = 0
                'Capture next Ticker's Open_Price
                Open_Price = ws.Cells(i + 1, 3).Value

                 'Reset the Stock Volume Total
                 'Yr_Yr_stock_vol_Total = 0
                
                'Hard Solution table calc on the current ws
                If (YearOver_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = YearOver_Percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (YearOver_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = YearOver_Percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
                'Hard Solution Reset
                YearOver_Percent = 0
                Total_Ticker_Volume = 0
             
            
             'Else - If the cell immediately following a row is still the same ticker name,
             'just add to Totl Ticker Volume
             Else
                'Encrease the Total Ticker Volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
             End If
            
      
         Next i

                'Hard Solution add values to table
                ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                ws.Range("P2").Value = MAX_TICKER_NAME
                ws.Range("P3").Value = MIN_TICKER_NAME
                ws.Range("Q4").Value = MAX_VOLUME
                ws.Range("P4").Value = MAX_VOLUME_TICKER

                'Hard Solution Summary Table - current worksheet
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"

                
        
     Next ws
     'I added this so no matter what sheet you run this on, you end up at the 1st sheet.
     Sheets(1).Select
     
     For Each sht In ThisWorkbook.Worksheets
            sht.Cells.EntireColumn.AutoFit
     Next sht
 
End Sub

'My code was created by many google searches, class reviews and youtube video's. I added some of my own ideas and touches, aka
'auto-sizing the summary table, selecting first sheet each time the code is run and so on.
'I found another way to do this database managemnent by using a 'Do Loop'. Cite: EverydayVBA from Youtube 'https://www.youtube.com/watch?v=piiDC4TQ1KE'
'I had much help from searching various github repos and finding
'https://freesoft.dev/program/163047389'

'As was mentioned in this Data Science programing opening, using google-fu is crucial in learning programing
'or programming in general. The truth is, the code is out there. Even though I was able to find the code
'while searching for help, I have gone through this code step by step to understand
'it's actions. This project was tough for me and I couldn't have done it without google-fu and all
'our in-class activites + office hours....





