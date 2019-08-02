Sub stocktckr():
    
    Dim YearChange As Double
    Dim Perchange As Double
    Dim Ticker_Location As Integer
    
    
    Dim stockvol As Double
    Dim stocktcker As String
    'set variable to total stock volume
    stockvol = 0
    
    
    For Each ws In Worksheets
        Dim lastrowticker As Double
            lastrowticker = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Ticker_Location = 2
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Volume"
                For i = 2 To lastrowticker
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        'set the ticker name and stockvol
    
                         Ticker_name = ws.Cells(i, 1).Value
                         stockvol = stockvol + ws.Cells(i, 7).Value
                         ws.Range("I" & Ticker_Location).Value = Ticker_name
                         ws.Range("J" & Ticker_Location).Value = stockvol
            
                         Ticker_Location = Ticker_Location + 1
            
        'If the next cell is the same brand..
                    Else
        
            'Add to vol total
                         stockvol = stockvol + ws.Cells(i, 7).Value
            
                    End If
               
                Next i
                    
       Next ws
        
End Sub
