
Sub Stock_Data()

    For Each ws In Worksheets
    
    Dim row_num As Integer
    Dim tot_vol As Double
    Dim open_val As Double
    Dim close_val As Double
    Dim count As Integer
        
    row_num = 2
    open_val = 0
    close_val = 0
    count = 0
    tot_vol = 0
    
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("K:K").NumberFormat = "0.00%"
    
    For i = 2 To lastrow
        
        count = count + 1
                
        'If statement checking if rows are the same
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
            'Adding the opening value
            open_val = open_val + ws.Cells(i, 3).Value
            
            'Adding the closing value
            close_val = close_val + ws.Cells(i, 6).Value
            
            'Adding the total volume
            tot_vol = tot_vol + ws.Cells(i, 7).Value
            
        'If statement checking value if rows are different
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Putting the ticker name in the ticker section of the outputs
            ws.Cells(row_num, 9).Value = ws.Cells(i, 1).Value
            
            'Adding the opening value
            open_val = open_val + ws.Cells(i, 3).Value
            
            'Adding the closing value
            close_val = close_val + ws.Cells(i, 6).Value
            
            'Adding the total volume
            tot_vol = tot_vol + ws.Cells(i, 7).Value
            
            'Outputs yearly change
            ws.Cells(row_num, 10).Value = (close_val - open_val)
            
                If close_val <> 0 Then
                
                'Outputs percentage change
                ws.Cells(row_num, 11).Value = (((close_val - open_val) / open_val) * 100)
                
                ElseIf close_val = 0 Then
                
                'Outputs percentage change
                ws.Cells(row_num, 11).Value = (((close_val - open_val) / 1) * 100)
                
                End If
            
            'Outputs total volume
            ws.Cells(row_num, 12).Value = tot_vol
                    
            'Increments row number for next row of output data
            row_num = row_num + 1
            
            'Reinitilizes variables for new sets of data
            open_val = 0
            close_val = 0
            tot_vol = 0
            count = 0
            
        End If
        
        'If statement to check value for conditional formatting
        If ws.Cells(i, 10).Value < 0 Then
        
        'Assigns the color red to values less than 0
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(i, 10).Value > 0 Then
            
            'Assigns the color green to values greater than 0
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If
            
    Next i
    
    Next ws
            
End Sub