Sub Magic_Button()
    
' Loop to go over all worksheets in the given excel file
For Each ws In Worksheets
    
' The part listed below computes the yearly change, percent change and total volume values for each stock
    Dim total As String
    total = 0
    
    Dim LastRow, op, cp, yc, pc As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' ws.Cells(2, 10).Value = LastRow ' test
    
    Dim sc, dc, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11 As Integer
    sc = 1    ' sc counts the number of stocks
    dc = 0    ' dc counts the numbef of days for each stock
    c1 = 1    ' column # for input ticker
    c2 = 7    ' column # for input volume
    c3 = 11   ' column # for output ticker
    c4 = 14   ' column # for output total volume
    c5 = 12   ' column # for yearly change output
    c6 = 13   ' column # for percent change output
    c7 = 3    ' column # for opening prices
    c8 = 6    ' column # for closing prices
    c9 = 16   ' column # for output greatest categories
    c10 = 17  ' column # for output ticker in "greatest" analysis part
    c11 = 18  ' column # for output greatest values
    
    ws.Cells(1, c3).Value = "Ticker"
    ws.Cells(1, c5).Value = "Yearly Change"
    ws.Cells(1, c6).Value = "Percent Change"
    ws.Cells(1, c4).Value = "Total Stock Volume"
        
    For i = 2 To LastRow
        If ws.Cells(i + 1, c1).Value = ws.Cells(i, c1).Value Then
            total = total + ws.Cells(i, c2).Value
            dc = dc + 1
            ' ws.Cells(i, 10).Value = dc ' test
        Else
            sc = sc + 1
            dc = dc + 1
            ' ws.Cells(i, 9).Value = dc ' test
            
            op = ws.Cells(i - dc + 1, c7).Value 'op:the first opening price of a stock
            cp = ws.Cells(i, c8).Value 'cp:the last closing price of a stock
            ' ws.Cells(1, 10).Value = op ' test
            ' ws.Cells(2, 10).Value = cp ' test

            yc = cp - op    ' yc:yearly change
            ws.Cells(sc, c5).Value = yc
                If yc > 0 Then
                    ws.Cells(sc, c5).Interior.ColorIndex = 4
                Else
                    ws.Cells(sc, c5).Interior.ColorIndex = 3
                End If
                
                If op > 0 Then
                    pc = (yc / op) ' pc:percent change
                    ws.Cells(sc, c6).Value = Format(pc, "Percent")
                Else
                    ws.Cells(sc, c6).Value = "" ' if op is 0 pc can't be calculated
                End If
                
            total = total + ws.Cells(i, c2).Value
            ws.Cells(sc, c4).Value = total
            ws.Cells(sc, c3).Value = ws.Cells(i, c1).Value
            
            ' reseting all values for the next stock
            total = 0
            dc = 0
            op = 0
            cp = 0
            yc = 0
            pc = 0
            
        End If
    
    Next i

' The part listed below finds out the highest and lowest percent changes and as well as the greatest total volume and the stocks associated with those
    
    Dim LastRow_for_columns_c4_c6, max_memory_pc, min_memory_pc, max_pc, min_pc, memory_tv, max_tv As Double
    LastRow_for_columns_c4_c6 = ws.Cells(Rows.Count, c5).End(xlUp).Row
        ' ws.Cells(1, 16).Value = LastRow_for_columns_c5_c6 ' test
    
    Dim max_pc_ticker, min_tv_ticker, max_tv_ticker As String

    max_memory_pc = ws.Cells(2, c6).Value
    min_memory_pc = ws.Cells(2, c6).Value
    max_pc_ticker = ws.Cells(2, c3).Value
    min_pc_ticker = ws.Cells(2, c3).Value
    
    For i = 3 To LastRow_for_columns_c4_c6
        If ws.Cells(i, c6).Value < max_memory_pc Then
            max_pc = max_memory_pc
        Else
            max_pc = ws.Cells(i, c6).Value
            max_memory_pc = ws.Cells(i, c6).Value
            max_pc_ticker = ws.Cells(i, c3).Value
        End If
        
        If ws.Cells(i, c6).Value < min_memory_pc Then
            min_pc = ws.Cells(i, c6).Value
            min_memory_pc = ws.Cells(i, c6).Value
            min_pc_ticker = ws.Cells(i, c3).Value
        Else
            min_pc = min_memory_pc
        End If
    Next i
                
    memory_tv = ws.Cells(2, c4).Value
    max_tv_ticker = ws.Cells(2, c3).Value
    
    For i = 3 To LastRow_for_columns_c4_c6
        If ws.Cells(i, c4).Value < memory_tv Then
            max_tv = memory_tv
        Else
            max_tv = ws.Cells(i, c4).Value
            memory_tv = ws.Cells(i, c4).Value
            max_tv_ticker = ws.Cells(i, c3).Value
        End If
    Next i
    

    ws.Cells(2, c9).Value = "Greatest % Increase"
    ws.Cells(3, c9).Value = "Greatest % Decrease"
    ws.Cells(4, c9).Value = "Greatest Total Volume"
    
    ws.Cells(1, c10).Value = "Ticker"
    ws.Cells(2, c10).Value = max_pc_ticker
    ws.Cells(3, c10).Value = min_pc_ticker
    ws.Cells(4, c10).Value = max_tv_ticker
    
    ws.Cells(1, c11).Value = "Value"
    ws.Cells(2, c11).Value = Format(max_pc, "Percent")
    ws.Cells(3, c11).Value = Format(min_pc, "Percent")
    ws.Cells(4, c11).Value = max_tv

Next ws

End Sub


