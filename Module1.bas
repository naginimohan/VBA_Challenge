Attribute VB_Name = "Module1"
Sub stocks_Peryear()

 For Each ws In Worksheets
 
        Dim WorksheetName As String
        Dim i As Long
        
        'Index counter to fill Ticker row
        Dim TickCount As Long
        Dim open_value As Double
        open_value = ws.Cells(2, 3).Value
        Dim close_value As Double
        Dim yearly_change As Double
        Dim total_stock_volume As Double
        total_stock_volume = 0
        
         WorksheetName = ws.Name
         'setting the headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total  stock volume"
         TickCount = 2
         
          'getting the last row  to loop
          lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
        
          For i = 2 To lastRow
              
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                  'Write ticker in column I (#9)
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                close_value = ws.Cells(i, 6).Value
                yearly_change = close_value - open_value
                'set value for yearly_change
               ws.Cells(TickCount, 10).Value = yearly_change
               'set value for percent_change
               If open_value = 0 Then
                ws.Cells(TickCount, 11).Value = 0
                Else
               ws.Cells(TickCount, 11).Value = yearly_change / open_value
                ws.Cells(TickCount, 11).NumberFormat = "0.00%"
               End If
               total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
               ws.Cells(TickCount, 12).Value = total_stock_volume
               'Reset the opening price
              open_value = ws.Cells(i + 1, 3)
                 TickCount = TickCount + 1
                 
               total_stock_volume = 0
            Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            ws.Cells(TickCount, 12).Value = total_stock_volume
        End If
          Next i
          
            lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
             'yearly change color formatting depending on positive and negative numbers
        For j = 2 To lastrow_summary_table
            
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
          For i = 2 To lastrow_summary_table
        
            'Find the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
          Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
        Next ws
End Sub
