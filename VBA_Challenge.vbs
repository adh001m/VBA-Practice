Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
	'Format the output sheet on All Stocks Analysis worksheet
	Worksheets("All Stocks Analysis Refactored").Activate
	
	Range("A1").Value = "All Stocks (" + yearvalue + ")"
	
	'Create a header row
	Cells(3, 1).Value = "Ticker"
	Cells(3, 2).Value = "Total Daily Volume"
	Cells(3, 3).Value = "Return"
	
	'Initialize array of all tickers
	Dim tickers(11) As String
	Dim tickerStartingPrices As Single
	Dim tickerEndingPrices As Single
	Dim tickerVolumes As Long
	
	tickers(0) = "AY"
	tickers(1) = "CSIQ"
	tickers(2) = "DQ"
	tickers(3) = "ENPH"
	tickers(4) = "FSLR"
	tickers(5) = "HASI"
	tickers(6) = "JKS"
	tickers(7) = "RUN"
	tickers(8) = "SEDG"
	tickers(9) = "SPWR"
	tickers(10) = "TERP"
	tickers(11) = "VSLR"
	
	'Activate data worksheet
	Worksheets("2018").Activate
	
	'Get the number of rows to loop over
	RowCount = Cells(Rows.Count, "A").End(xlUp).Row
	
	
	''2a) Create a for loop to initialize the tickerVolumes to zero.
	For I = 0 To 11
	Ticker = tickers(I)
	tickerVolumes = 0
	    
	''2b) Loop over all the rows in the spreadsheet.
	    Worksheets("2018").Activate
	
	    '3a) Increase volume for current ticker
	    For j = 2 To RowCount
	    
	    If Cells(j, 1).Value = Ticker Then
	
	           tickerVolumes = tickerVolumes + Cells(j, 8).Value
	    End If
	    
	    
	    
	    
	    '3b) Check if the current row is the first row with the selected tickerIndex.
	    'If  Then
	        
	          If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
	
	           tickerStartingPrices = Cells(j, 6).Value
	        
	    End If
	    
	    '3c) check if the current row is the last row with the selected ticker
	     'If the next row’s ticker doesn’t match, increase the tickerIndex.
	    'If  Then
	        
	        
	        If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then tickerEndingPrices = Cells(j, 6).Value
	       
	
	 Next j
	        
	
	Next I
	
	'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
	For I = 0 To 11
	    
	    Worksheets("All Stocks Analysis Refactored").Activate
	    
	   Cells(4 + I, 1).Value = Ticker
	   Cells(4 + I, 2).Value = tickerVolumes
	   Cells(4 + I, 3).Value = (tickerEndingPrices / tickerStartingPrices) - 1
	
	
	Next I
	
	'Formatting
	Worksheets("All Stocks Analysis Refactored").Activate
	Range("A3:C3").Font.FontStyle = "Bold"
	Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
	Range("B4:B15").NumberFormat = "#,##0"
	Range("C4:C15").NumberFormat = "0.0%"
	Columns("B").AutoFit
	
	dataRowStart = 4
	dataRowEnd = 15
	
	For I = dataRowStart To dataRowEnd
	    
	    If Cells(I, 3) > 0 Then
	        
	        Cells(I, 3).Interior.Color = vbGreen
	        
	    Else
	    
	        Cells(I, 3).Interior.Color = vbRed
	        
	    End If
	    
	Next I
	
	endTime = Timer
	MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)
End Sub