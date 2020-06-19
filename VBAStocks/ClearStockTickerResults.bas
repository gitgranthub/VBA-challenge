Attribute VB_Name = "CLEAR_ALL_Module"
Sub ButtonClearStockResults()

Dim ws As Worksheet
    
    'Begin the loop to clear solved cells/columns.
    For Each ws In Worksheets
                
                'activate current ws in loop
                ws.Activate
                
                ws.Range("I2:I" & Range("I2").End(xlDown).Row).Clear
                ws.Range("J2:J" & Range("J2").End(xlDown).Row).Clear
                ws.Range("K2:K" & Range("K2").End(xlDown).Row).Clear
                ws.Range("L2:L" & Range("L2").End(xlDown).Row).Clear
                
                'I use this clear action when I added the hard solve.
                ws.Range("O1:O" & Range("O1").End(xlDown).Row).Clear
                ws.Range("O1:O" & Range("O1").End(xlDown).Row).Clear
                ws.Range("O3:O" & Range("O3").End(xlDown).Row).Clear
                ws.Range("P1:P" & Range("P1").End(xlDown).Row).Clear
                ws.Range("Q1:Q" & Range("Q1").End(xlDown).Row).Clear



         Next
         'I thought this would be a nice touch. It takes you back to the 1st sheet no matter what sheet you run the macro on.
         Sheets(1).Select

'a little validation that the sheets are clear
MsgBox "All Clear!" + " Let's Check the results again"

End Sub

