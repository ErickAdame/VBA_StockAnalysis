# VBA_StockAnalysis

Sub ticker_analysis()

    Dim stocktotal As Double
    Dim rowindex As Long
    Dim change As Double 'where change is stock price change
    Dim tablerow As Integer
    Dim start As Long
    Dim rowcount As Long
    Dim percentchange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets 'Loop through each worksheet in the workbook
    
    tablerow = 0
    stocktotal = 0
    change = 0
    start = 2
    dailychange = 0
    
    'Create all the labels for the data on the table
    ws.Range("I1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    'Calculate the row count for each worksheet - confirmed with msgbox
    rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'time to loop through every row
For rowindex = 2 To rowcount
    
        If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
        
            stocktotal = stocktotal + ws.Cells(rowindex, 7).Value
            
            If stocktotal = 0 Then
                'print results
                ws.Range("I" & 2 + tablerow).Value = Cells(rowindex, 1).Value
                ws.Range("J" & 2 + tablerow).Value = 0
                ws.Range("K" & 2 + tablerow).Value = "%" And 0
                ws.Range("L" & 2 + tablerow).Value = 0
            Else
        'find the cell with the open stock price to calculate the change in stock
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To rowindex
                        If ws.Cells(find_value, 3).Value <> 0 Then
                        start = find_value  'establishes the row index with the open price
                        Exit For
                        End If
                    Next find_value
                End If
            
            'Calculate the stock change and percentage change
                change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
                percentchange = change / ws.Cells(start, 3)
                
                start = rowindex + 1
                
            'print the final calculations to the table
                ws.Range("I" & 2 + tablerow) = ws.Cells(rowindex, 1).Value
                ws.Range("j" & 2 + tablerow) = change
                ws.Range("j" & 2 + tablerow).NumberFormat = "0.00"
                ws.Range("K" & 2 + tablerow).Value = percentchange
                ws.Range("K" & 2 + tablerow).NumberFormat = "0.00%"
                ws.Range("L" & 2 + tablerow).Value = stocktotal
                
            'Conditional formatting for the table column with stock change... Red, Green or None
                Select Case change
                    Case Is > 0
                        ws.Range("J" & 2 + tablerow).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + tablerow).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + tablerow).Interior.ColorIndex = 0
                End Select
                
                
            
            End If
        'Reset data before looping again
        stocktotal = 0
        change = 0
        tablerow = tablerow + 1
        
        Else
        'Add up all the stock total volume for the table - variable is store under IF statement
        stocktotal = stocktotal + ws.Cells(rowindex, 7).Value
        
        End If
    
    Next rowindex
    
    Next ws
    

End Sub
