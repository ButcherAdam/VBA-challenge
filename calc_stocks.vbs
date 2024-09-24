Sub calc_stocks()

Dim ws As Worksheet

For Each ws In Worksheets

    'Variables for iterating
    Dim stock_index As Integer
    Dim stock_length As Long
    Dim stock_total As LongLong
    
    'Variables to calculate opening and closing prices
    Dim stock_open As Double
    Dim stock_close As Double
    
    'Initialize iterating variables
    stock_index = 2
    stock_length = ws.Range("A" & Rows.Count).End(xlUp).Row
    stock_total = 0
    
    'Obtain the first stock opening value
    stock_open = ws.Cells(2, 3).Value
    
    'Variables for the calculated headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    'Go through stocks
    For i = 2 To stock_length
    stock_total = stock_total + ws.Cells(i, 7).Value
    
    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) And i <> stock_length Then
        'Retrieve the Stock Name
        ws.Cells(stock_index, 9).Value = ws.Cells(i, 1).Value
        
        'Calculate the Quarterly and Percent Changes
        stock_close = ws.Cells(i, 6).Value
        ws.Cells(stock_index, 10) = (stock_close - stock_open)
        If (ws.Cells(stock_index, 10).Value > 0) Then
            ws.Cells(stock_index, 10).Interior.ColorIndex = 4
        ElseIf (ws.Cells(stock_index, 10).Value < 0) Then
            ws.Cells(stock_index, 10).Interior.ColorIndex = 3
        End If
        
        ws.Cells(stock_index, 11) = ((stock_close - stock_open) / stock_open)
        ws.Cells(stock_index, 12).Value = stock_total
        
        'Move to the next Stock Name
        stock_open = ws.Cells(i + 1, 3)
        stock_total = 0
        stock_index = stock_index + 1
    End If
    
    'Go here when i reaches the last row
    If i = stock_length Then
        'Retrieve the Stock Name
        ws.Cells(stock_index, 9).Value = ws.Cells(i, 1).Value
        
        'Calculate the Quarterly and Percent Changes
        stock_close = ws.Cells(i, 6).Value
        ws.Cells(stock_index, 10) = (stock_close - stock_open)
        ws.Cells(stock_index, 11) = ((stock_close - stock_open) / stock_open)
        ws.Cells(stock_index, 12).Value = stock_total
    End If
    
    Next i

    'Variables for the summary headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    'Variables for the summary rows
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Variable to get information about the values calculated
    stock_name = ws.Range("K" & Rows.Count).End(xlUp).Row

    'Variables for calculating summary values
    Dim max_percent_increase As Double
    Dim max_percent_increase_index As Integer
    Dim min_percent_decrease As Double
    Dim min_percent_decrease_index As Integer
    Dim max_volume As LongLong
    Dim max_volume_index As Integer

    
    max_percent_increase = 0
    max_percent_increase_index = 0
    min_percent_decrease = 0
    min_percent_decrease_index = 0
    max_volume = 0
    max_volume_index = 0
    
    For j = 2 To stock_name
    'Format each percentage change cell to percentage
        ws.Range("K" & j).NumberFormat = "0.00%"
    
    'Compare max percentage increase
    If ws.Cells(j, 11).Value > max_percent_increase Then
    max_percent_increase = ws.Cells(j, 11).Value
    max_percent_increase_index = j
    End If
    
    'Compare min percentage decrease
    If ws.Cells(j, 11).Value < min_percent_decrease Then
    min_percent_decrease = ws.Cells(j, 11).Value
    min_percent_decrease_index = j
    End If

    'Compare max volume
    If ws.Cells(j, 12).Value > max_volume Then
    max_volume = ws.Cells(j, 12).Value
    max_volume_index = j
    End If
    Next j
    
    'Include summary values
    ws.Cells(2, 16).Value = ws.Cells(max_percent_increase_index, 9).Value
    ws.Cells(2, 17).Value = max_percent_increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = ws.Cells(min_percent_decrease_index, 9).Value
    ws.Cells(3, 17).Value = min_percent_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = ws.Cells(max_volume_index, 9).Value
    ws.Cells(4, 17).Value = max_volume
    
Next ws

End Sub