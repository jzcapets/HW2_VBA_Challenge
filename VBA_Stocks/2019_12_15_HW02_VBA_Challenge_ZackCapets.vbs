Attribute VB_Name = "Module1"
Sub stocks():
    
    'declare stock variables
    Dim tickername As String
    Dim tickertable As Double
    Dim tickervolume As Double
    Dim tickeropenprice As Double
    Dim tickercloseprice As Double
    Dim firstrowofticker As Double
    Dim lastrowofticker As Double
    Dim tickerpricechange As Double
    Dim tickerpctchange As Double
    Dim lastrow As Long
    
    'declare variables that track info needed for HW bonus
    Dim biggain As Double
    Dim bigloss As Double
    Dim highvol As Double
    Dim winner As String
    Dim loser As String
    Dim mover As String

    'declare iterator / length variables. i loops through all rows in a sheet, j loops through all rows in the summary table
    Dim i As Long
    Dim j As Long
    
    'loop through the worksheets
    For Each ws In Worksheets
    
        'get the worksheet names
        Dim WorksheetName As String
        WorksheetName = ws.name
        
        'find the last nonblank row in each sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'reset tickertable counter, this sets the first row that our ticker data will appear in
        tickertable = 1
        
        'make some column headers for our list of stock data summed by ticker symbol
        ws.Range("P" & tickertable).Value = "Ticker Name"
        ws.Range("Q" & tickertable).Value = "Intra-year price change ($)"
        ws.Range("R" & tickertable).Value = "Intra-year % price change (%)"
        ws.Range("S" & tickertable).Value = "Ticker Volume"
       
        'set first ticker open price
        tickeropenprice = ws.Cells(2, 3).Value
        'give firstrowofticker an initial value for first ticker
        firstrowofticker = 2
       
        'start on the 2nd row of each sheet since that's where the data begins
        For i = 2 To lastrow
        
        
        'if the ticker doesn't match the next ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                  ' Set the ticker name
                tickername = ws.Cells(i, 1).Value
                
                ' Push ticker name to the Summary Table
                ws.Range("P" & tickertable + 1).Value = tickername
                
                ' include volume from last ticker transaction in tickervolume being tallied in the else portion of this IF statement
                tickervolume = tickervolume + ws.Cells(i, 7).Value
                
                ' Print the Ticker Volume to the Summary Table
                ws.Range("S" & tickertable + 1).Value = Format(tickervolume, Scientific)
                             
                ' Set the ticker close price
                tickercloseprice = ws.Cells(i, 6).Value
                
                'Retrieve ticker open price
                tickeropenprice = ws.Cells(firstrowofticker, 3).Value
                
                    'In case a ticker has 0 volume or 0 price data values then the price change and percent change will overflow due to div/0
                    If tickeropenprice = 0 Then
                        tickerpctchange = 0
                    Else
                        'Calculate metrics for ticker
                        tickerpricechange = tickercloseprice - tickeropenprice
                        tickerpctchange = (tickerpricechange / tickeropenprice) * 100
                    End If
                  
                  'Put ticker intra-year price change in the Summary Table
                  ws.Range("Q" & tickertable + 1).Value = tickerpricechange
                  
                  'Put ticker intra-year percent change in price in the Summary Table
                  ws.Range("R" & tickertable + 1).Value = FormatPercent(tickerpctchange, 2)
                  
                    'fill price change cell red if price change less than zero, green if greater than 0
                    If (tickerpctchange < 0) Then
                          ws.Range("Q" & tickertable + 1).Interior.ColorIndex = 3
                      Else
                          ws.Range("Q" & tickertable + 1).Interior.ColorIndex = 4
                    End If
                    
                   'reset first row of ticker for next ticker
                  firstrowofticker = i + 1
                    
                  'Increment tickerrow
                  tickertable = tickertable + 1
                  
                  'reset ticker volume to begin counting next ticker's volume
                  tickervolume = 0
               
               'do this if ticker(i) = ticker(i+1)
                Else
                                
                  'keep counting ticker volume while ticker is the same as previous
                  tickervolume = tickervolume + ws.Cells(i, 7).Value
            End If
        
        'next row
        Next i
    
        'now loop through the summary table row by row and compare each row's values against our variables that hold our extremes. Each time we find a new extreme that value gets pushed to our extreme variables
        
        'give the variables that hold our extreme variables an initial value to compare against
        biggain = 0
        bigloss = 0
        highvol = 0
        
        
        For j = 2 To tickertable
                name = ws.Cells(j, 16).Value
                'winners/loser loop
                If (ws.Cells(j, 18) > biggain) Then
                    biggain = ws.Cells(j, 18).Value
                    winner = name
                ElseIf ws.Cells(j, 18) < bigloss Then
                    bigloss = ws.Cells(j, 18).Value
                    loser = name
                End If
                
                'moversloop
                If ws.Cells(j, 19) > highvol Then
                    highvol = ws.Cells(j, 19).Value
                    mover = name
                End If
        Next j
        
        'make some row headers on the summary sheet
        ws.Range("U1").Interior.ColorIndex = 1
        ws.Range("U2").Value = "Ticker:"
        ws.Range("U3").Value = "Value:"
        
        'make some column headers on the summary sheet
        ws.Range("V1").Value = "Greatest % increase:"
        ws.Range("W1").Value = "Greatest % decrease:"
        ws.Range("X1").Value = "Greatest total volume:"
        
        'fill in the tickers
        ws.Range("V2").Value = winner
        ws.Range("W2").Value = loser
        ws.Range("X2").Value = mover
        
        'fill in the tickers' values
        ws.Range("V3").Value = FormatPercent(biggain, 2)
        ws.Range("W3").Value = FormatPercent(bigloss, 2)
        ws.Range("X3").Value = highvol
    
    'nextworksheet
    Next ws
    
End Sub
