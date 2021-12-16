Attribute VB_Name = "Module1"
Sub Summarize()

'Part 1
For Each ws In Worksheets

    'declaring variables
    'Not sure why I had to make variables single.
    'Defining as double gave me rounding errors.
    Dim startprice As Single
    Dim endprice As Single
    Dim yearlychange As Single
    Dim percentchange As Single
    Dim totalvolume As Double
    Dim ticker As String
    Dim summaryrow As Integer
    
    'labeling the summary table headers
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total stock volume"
    'ws.Range("N1").Value = "start price"
    'ws.Range("O1").Value = "end price"
    
    'starting the summary row at 2
    summaryrow = 2
    'determining maximum value for for loop
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    'captures the starting price
    startprice = ws.Cells(2, 3).Value
     
    For i = 2 To lastrow
        'summation for totalstock value
        totalvolume = totalvolume + ws.Cells(i, 7).Value
            
        'if next row ticker symbol is different from current row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'it will log and report the current ticker symbol
            ticker = ws.Cells(i, 1).Value
            ws.Cells(summaryrow, 9).Value = ticker
            
            'it will log the end value
            endprice = ws.Cells(i, 6).Value
            
            'checking if the printing is correct
            'ws.Cells(summaryrow, 14).Value = startprice
            'ws.Cells(summaryrow, 15).Value = endprice
            
            'it will report yearly change
            yearlychange = endprice - startprice
            ws.Cells(summaryrow, 10).Value = yearlychange
           
            
            'it will report percent change if divider is not 0
            If startprice <> 0 Then
                percentchange = (yearlychange / startprice)
                ws.Cells(summaryrow, 11).Value = percentchange
                
                
            
            End If
            
            'color conditional formatting
                If ws.Cells(summaryrow, 10).Value = 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 2
                
                ElseIf ws.Cells(summaryrow, 10).Value > 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(summaryrow, 10).Value < 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
            
            End If
            
            'it will report total stock volume
            ws.Cells(summaryrow, 12).Value = totalvolume
                        
            'move to next summaryrow
            summaryrow = summaryrow + 1
            'reset total volume
            totalvolume = 0
            'logs next ticker start price
            startprice = ws.Cells(i + 1, 3).Value
            
          End If
    Next i
 
 ws.Range("J:J").NumberFormat = "0.00"
 ws.Range("K:K").NumberFormat = "0.00%"
        
Next


'Part 2

For Each ws In Worksheets
    
    'declaring variables
    Dim increaseindex As Integer
    Dim decreaseindex As Integer
    Dim volumeindex As Integer
    
    increaseindex = 2
    decreaseindex = 2
    volumeindex = 2
    
    'labeling the analysis  headers
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "ticker"
    ws.Range("P1").Value = "value"
    
    'getting the row count for # of unique tickers
    finalrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'scanning summary table
    For i = 2 To finalrow
        
        'check for greatest % increase
        If ws.Cells(increaseindex, 11).Value < ws.Cells(i, 11).Value Then
            increaseindex = i
        End If
                
        'check for greatest % decrease
        If ws.Cells(decreaseindex, 11).Value > ws.Cells(i, 11).Value Then
            decreaseindex = i
        End If
        
        'check for greatest total volume
        If ws.Cells(volumeindex, 12).Value < ws.Cells(i, 12).Value Then
            volumeindex = i
        End If
    
    Next i
    
    'reporting analysis table
    'greatest % increase
    'ticker
    ws.Range("O2").Value = ws.Cells(increaseindex, 9).Value
    'value
    ws.Range("P2").Value = ws.Cells(increaseindex, 11).Value
    
    'greatest % decrease
    'ticker
    ws.Range("O3").Value = ws.Cells(decreaseindex, 9).Value
    'value
    ws.Range("P3").Value = ws.Cells(decreaseindex, 11).Value
    
    'greatest volume
    'ticker
    ws.Range("O4").Value = ws.Cells(volumeindex, 9).Value
    'value
    ws.Range("P4").Value = ws.Cells(volumeindex, 12).Value

ws.Range("P2:P3").NumberFormat = "0.00%"
 

Next



End Sub
