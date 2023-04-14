Attribute VB_Name = "Module1"
Sub YearlyStockAnalysis():

    For Each ws In Worksheets
    
        Dim d As Long
        Dim c As Long
        Dim StockAnalysis As String
        Dim Counter As Long
        Dim RowA As Long
        Dim RowI As Long
        Dim PercentChange As Double
        Dim Increase As Double
        Dim Decrease As Double
        Dim Volume As Double
        
        'Get the WorksheetName
        StockAnalysis = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        Counter = 2
        
        'Set start row to 2
        c = 2
        
        'Locate last occupied cell in column A
        RowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop
            For d = 2 To RowA
            
                If ws.Cells(d + 1, 1).Value <> ws.Cells(d, 1).Value Then
                
                'Write "Ticker" in (1, I)
                ws.Cells(Counter, 9).Value = ws.Cells(d, 1).Value
                
                'Write "Yearly Change" in (1, J) and find values
                ws.Cells(Counter, 10).Value = ws.Cells(d, 6).Value - ws.Cells(c, 3).Value
                
                    
                    If ws.Cells(Counter, 10).Value < 0 Then
                
                    'Set color to red
                    ws.Cells(Counter, 10).Interior.ColorIndex = 9
                
                    Else
                
                    'Set all else to green
                    ws.Cells(Counter, 10).Interior.ColorIndex = 10
                
                    End If
                    
                    'Write "Percent Change" in (1, K)
                    If ws.Cells(c, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(d, 6).Value - ws.Cells(c, 3).Value) / ws.Cells(c, 3).Value)
                    
                    'Percent formating
                    ws.Cells(Counter, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Counter, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Write "Total Volume" in (1, L)
                ws.Cells(Counter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(c, 7), ws.Cells(d, 7)))
                
                'Increase Counter by 1
                Counter = Counter + 1
                
                c = d + 1
                
                End If
            
            Next d
            
        'Locate last occupied row for column I
        RowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Summary
        Volume = ws.Cells(2, 12).Value
        Increase = ws.Cells(2, 11).Value
        Decrease = ws.Cells(2, 11).Value
        
            'Loop
            For d = 2 To RowI
            
                'For volume, if next value is larger then replace and populate cell
                If ws.Cells(d, 12).Value > Volume Then
                Volume = ws.Cells(d, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(d, 9).Value
                
                Else
                
                Volume = Volume
                
                End If
                
                'For increase, if next value is larger then replace and populate cell
                If ws.Cells(d, 11).Value > Increase Then
                Increase = ws.Cells(d, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(d, 9).Value
                
                Else
                
                Increase = Increase
                
                End If
                
                'For decrease, if next value is larger then replace and populate cell
                If ws.Cells(d, 11).Value < Decrease Then
                Decrease = ws.Cells(d, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(d, 9).Value
                
                Else
                
                Decrease = Decrease
                
                End If
                
            'Display results summary
            ws.Cells(2, 17).Value = Format(Increase, "Percent")
            ws.Cells(3, 17).Value = Format(Decrease, "Percent")
            ws.Cells(4, 17).Value = Format(Volume, "Scientific")
            
            Next d
            
        'Column width
        Worksheets(StockAnalysis).Columns("A:R").AutoFit
            
    Next ws
        
End Sub
