Sub ButtonClearStockResults()

'****MAKE THIS A BUTTON'****
'this clears data from columns (I,J,K,L)'
'Run before you rerun the program'

'***?USe this if you want to clear a specific Worksheet column or columns***'
'Worksheets(sheetname).Columns(1).ClearContents '

'these clear specific columns (we specify) in the current worksheet'
ActiveSheet.Range("I2:I" & Range("I2").End(xlDown).Row).ClearContents
ActiveSheet.Range("J2:J" & Range("J2").End(xlDown).Row).ClearContents
ActiveSheet.Range("K2:I" & Range("K2").End(xlDown).Row).ClearContents
ActiveSheet.Range("L2:I" & Range("L2").End(xlDown).Row).ClearContents

MsgBox "All Clear!" + " Let's Check the results again"


End Sub