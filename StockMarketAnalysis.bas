Attribute VB_Name = "Module1"
'loop all the cells in column 1
'if the ticker is the same, keep grapping the trade volumn and add them up to put down into cell (N2)
'grap the opening price at the beginning of the year
'at the end of the year, grap the closing price. Then do the yearly change
'use that result to calculate the percentage change


Sub stockMarket()
    
    'grab the cell, and count how many cells to that last one
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'create defaul variables to use as counter/container. So we can put them into the summary table.
    Dim ticker As String
    
    Dim ticker_counter As Integer
    ticker_counter = 0
    
    Dim yearly_change As Double
    
    
    'Dim total_volume As Integer
    total_volume = 0
    
    'create a variable that hold the year opening price
    Dim year_opening As Double
    year_opening = 0
    
    'create a variable that hold the year closing price
    Dim year_closing As Double
    year_closing = 0
    
    
    'create a varible to keep track of summary table position
    summary_table_row = 2
    
    
    For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'hold the ticker letter into the variable
            ticker = Cells(i, 1).Value
            'print the ticker on the K column
            Cells(summary_table_row, 11).Value = ticker
            
            'grab the closing price at the end of the year
            year_closing = Cells(i, 6).Value
            year_opening = Cells(i - ticker_counter, 3).Value
            
            'print the yearly change into column L
            yearly_change = year_closing - year_opening
            Cells(summary_table_row, 12).Value = yearly_change
            'print the percentage change on the column M
            Cells(summary_table_row, 13).Value = (yearly_change / year_opening)
            
            'coloring the cells in column L, red for negative, green for positive or equal to 0
            If yearly_change >= 0 Then
                Cells(summary_table_row, 12).Interior.ColorIndex = 4
            Else
                Cells(summary_table_row, 12).Interior.ColorIndex = 3
            End If
            
            'update the total trading volumn
            total_volume = total_volume + Cells(i, 7).Value
            
            'print the total trading volumn into the N column
            Cells(summary_table_row, 14).Value = total_volume
            
            'update the summary table row to move to the next line
            summary_table_row = summary_table_row + 1
            
            'reset the trading volume to 0, to count the new ticker
            total_volume = 0
            
            'set ticker counter back to 0
            ticker_counter = 0
        Else
            total_volume = total_volume + Cells(i, 7).Value
            ticker_counter = ticker_counter + 1
        End If
    Next i
    
End Sub

