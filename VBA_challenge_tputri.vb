Attribute VB_Name = "alphabetical_testing.xlsx"

Sub Summarize()

    Dim ticker As String
    Dim yearlychange As Single
    Dim percentchange As Single
    Dim totalstockvolume As Double
    Dim greatest$increase As integer
    Dim greatest$decrease As integer
    Dim greatesttotalvolume as integer  
    Dim summaryrow As Integer
    Dim totalvolume As Double
   
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total stock volume"
    
    'starting at row no 2
    summaryrow = 2
    'max value for the loop
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'capture starting price
    startprice = ws.Cells(2, 3).Value

    For i = 2 To lastrow
        totalvolume = totalvolume + ws.Cells(i, 7).Value

        'if next row ticker symbol is different from current row
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ws.Cells(summaryrow, 9).Value = ticker
            endprice = ws.Cells(i, 6).Value

            yearlychange = endprice - startprice
            ws.Cells(summaryrow, 10).Value = yearlychange

        If startprice <> 0 Then
            percentchange = (yearlychange / startprice)
            ws.Cells(summaryrow, 11).Value = percentchange

        end If

            If ws.Cells(summaryrow, 10).Value = 0 Then
            ws.Cells(summaryrow, 10).Interior.ColorIndex = 2
                
            ElseIf ws.Cells(summaryrow, 10).Value > 0 Then
            ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(summaryrow, 10).Value < 0 Then
            ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
            
        End If

        ws.Cells(summaryrow, 12).Value = totalvolume
        summaryrow = summaryrow + 1
        totalvolume = 0
        startprice = ws.Cells(i + 1, 3).Value

        End If

    Next i
 
 ws.Range("J:J").NumberFormat = "0.00"
 ws.Range("K:K").NumberFormat = "0.00%"
        
Next

For Each ws In Worksheets
	    
	    
	Dim increaseindex As Integer
	Dim decreaseindex As Integer
	Dim volumeindex As Integer
	    
	increaseindex = 2
	decreaseindex = 2
	volumeindex = 2
	    
	ws.Range("N2").Value = "Greatest % Increase"
	ws.Range("N3").Value = "Greatest % Decrease"
	ws.Range("N4").Value = "Greatest Total Volume"
	ws.Range("O1").Value = "ticker"
    ws.Range("P1").Value = "value"
	    
	finalrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
	    
	
	For i = 2 To finalrow
	        
	        'for greatest % increase
	    If ws.Cells(increaseindex, 11).Value < ws.Cells(i, 11).Value Then
	        increaseindex = i
	    End If
	                
	        'for greatest % decrease
	    If ws.Cells(decreaseindex, 11).Value > ws.Cells(i, 11).Value Then
	        decreaseindex = i
	    End If
	        
	        ' for greatest total volume
	    If ws.Cells(volumeindex, 12).Value < ws.Cells(i, 12).Value Then
	        volumeindex = i
	    End If
	    
	    Next i
	    
	    ws.Range("O2").Value = ws.Cells(increaseindex, 9).Value
	    
	    ws.Range("P2").Value = ws.Cells(increaseindex, 11).Value
	    
	    ws.Range("O3").Value = ws.Cells(decreaseindex, 9).Value
	    
	    ws.Range("P3").Value = ws.Cells(decreaseindex, 11).Value
	    
	    ws.Range("O4").Value = ws.Cells(volumeindex, 9).Value
	   
	    ws.Range("P4").Value = ws.Cells(volumeindex, 12).Value
	
	    ws.Range("P2:P3").NumberFormat = "0.00%"
	 
	
	Next
	

	










