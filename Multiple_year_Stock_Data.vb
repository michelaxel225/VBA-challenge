Sub Multipe_Year_Stock_Data()
  'Loop Through all the sheets 
   For Each ws In Worksheets
    Dim WorksheetName As String
    WorksheetName = ws.Name
   ' Add Ticker , Yearly change, Percent Change and Total stock volume as column header 
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
     ws.Cells(1, "L").Value = "Total Stock Volume"

       'Define variables 
        Dim Ticker_Name As String
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
         Dim Percent_Change As Double
        Dim Total_stock_Volume As Double
        Total_stock_Volume = 0 
         Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
          Dim i As Long
     'Determine the last  row
          lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          For i = 2 To lastrow
         
         ' Loop through all ticker symbol
        
      If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                ws.Cells(Row, Column + 8).Value = Ticker_Name
                
                'Set Initial Opening Price
                Opening_Price = ws.Cells(2, Column + 2).Value
                ' Set Closing Price
                Closing_Price = ws.Cells(i, Column + 5).Value
                ' Set Yearly Change
                Yearly_Change = Closing_Price - Opening_Price
                ws.Cells(Row, Column + 9).Value = Yearly_Change
                'Set Percent Change
                If (Opening_Price = 0 And Closing_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Opening_Price
                    ws.Cells(Row, Column + 10).Value = Percent_Change
                  'format the percentage 
                    ws.Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total stock volume 
                Total_stock_Volume = Total_stock_Volume + ws.Cells(i, Column + 6).Value
                ws.Cells(Row, Column + 11).Value = Total_stock_Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset Opening Price
                Opening_Price = ws.Cells(i + 1, Column + 2)
                ' reset the total stock volume 
                Total_stock_Volume = 0
            
           else
                Total_stock_Volume = Total_stock_Volume + ws.Cells(i, Column + 6).Value
          end if 
           ws.Cells(Row, Column + 11).NumberFormat = "0"
        Next i
        
        ' Determine the Last Row of Yearly Change in order to set cell colors 
        Yearly_Change_LastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        'set the colors 
        For j = 2 To Yearly_Chanage_LastRow
            If (Cells(j, 10).Value >0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    
Next ws 
End Sub
