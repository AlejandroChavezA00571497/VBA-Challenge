Attribute VB_Name = "Module1"




Sub StockAnalysis01()

'I declare my variables
Dim Ticker, GPIncTicker, GPDecTicker, GTVolTicker As String
Dim YChange, PChange, TotalVol, FirstRow, LastRow, LastRow2, TickerIndex, i, j, GPInc, GPDec, GTVol As Double
Dim OPrice, CPrice As Double

'I start the loop that goes through all Worksheets

For Each ws In Worksheets

    'I initialize values for the ticker index, volume and opening price, as well as the greatest increases and volume
    TickerIndex = 2
    TotalVol = 0
    OPrice = ws.Cells(2, 3).Value
    GPInc = 0
    GPDec = 0
    GTVol = 0
    
    'I assign the last row value to its variable
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'I create the headers for the result columns in all worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'I open the loop that goes through all of the rows in the ws
    For i = 3 To LastRow
        
        'I check that we are still on the same ticker
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i - 1, 1).Value
            
            'I get the closing price for the ticker we just closed
            CPrice = ws.Cells(i - 1, 6).Value
            
            'I get the yearly change and percentage change
            YChange = CPrice - OPrice
            
            If OPrice <> 0 Then
                PChange = (YChange / OPrice)
            Else
                PChange = 0
            End If
            
            'I populate the cells with that data
            ws.Range("I" & TickerIndex).Value = Ticker
            ws.Range("J" & TickerIndex).Value = YChange
            ws.Range("K" & TickerIndex).Value = PChange
            ws.Range("L" & TickerIndex).Value = TotalVol
            
            
            'I move to the next row in the results section, and reset the total volume
            TickerIndex = TickerIndex + 1
            TotalVol = 0
            
            'I get the opening price for the new ticker
            OPrice = ws.Cells(TickerIndex, 3).Value
        
        Else
            
            'If we are still on the same ticker, I sum the value of the total volume
            TotalVol = TotalVol + ws.Cells(TickerIndex, 7).Value
        
        
        End If
        
        'I check to see if there is a new greatest increase or decrease in percentage as well as total volume
        If PChange > GPInc Then
        
            GPIncTicker = Ticker
            GPInc = PChange
            
            ElseIf PChange < GPDec Then
            
                GPDecTicker = Ticker
                GPDec = PChange
            
            End If
            
            If TotalVol > GTVol Then
            
                GTVolTicker = ws.Cells(i - 1, 1).Value
                GTVol = TotalVol
            
            End If
            
        
        
        
    Next i
    
    
    
    
    
    
    'I format the cells for the results that they'll be carrying
    ws.Range("K:K").NumberFormat = "0.00%"
    
    'I add conditional formatting for the yearly change, if negative the cell will be red, if positive, green
    LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row
    For j = 2 To LastRow2
    
        If ws.Cells(j, 10).Value <= 0 Then
        
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
            ElseIf ws.Cells(j, 10).Value > 0 Then
            
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
            End If
        
    Next j
     
    

    'I populate the results of greatest percentage increase, decrease and total volume
    
    ws.Cells(2, 16).Value = GPIncTicker
    ws.Cells(3, 16).Value = GPDecTicker
    ws.Cells(4, 16).Value = GTVolTicker
    
    ws.Cells(2, 17).Value = FormatPercent(GPInc)
    ws.Cells(3, 17).Value = FormatPercent(GPDec)
    ws.Cells(4, 17).Value = FormatNumber(GTVol)
    

            
Next ws

End Sub


Sub ClearCells()


'I also made a script to clear the values from all worksheets
For Each ws In Worksheets

ws.Range("I1:L999999").ClearContents
ws.Range("I1:L999999").Interior.ColorIndex = 2
ws.Range("P2:Q4").ClearContents

Next ws

End Sub





