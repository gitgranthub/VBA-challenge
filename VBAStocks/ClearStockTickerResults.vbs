Sub ButtonClearStockResults()

Dim ws As Worksheet
    
    'Begin the loop.
    For Each ws In Worksheets
                ws.Activate
                
                ws.Range("I2:I" & Range("I2").End(xlDown).Row).Clear
                ws.Range("J2:J" & Range("J2").End(xlDown).Row).Clear
                ws.Range("K2:K" & Range("K2").End(xlDown).Row).Clear
                ws.Range("L2:L" & Range("L2").End(xlDown).Row).Clear
                ws.Range("O1:O" & Range("O1").End(xlDown).Row).Clear
                ws.Range("O3:O" & Range("O3").End(xlDown).Row).Clear
                ws.Range("P1:P" & Range("P1").End(xlDown).Row).Clear
                ws.Range("Q1:Q" & Range("Q1").End(xlDown).Row).Clear



         Next
         
         Sheets(1).Select

MsgBox "All Clear!" + " Let's Check the results again"

End Sub
