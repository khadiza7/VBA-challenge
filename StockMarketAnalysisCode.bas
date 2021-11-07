Attribute VB_Name = "Module1"
Sub StockMarketAnalysis():


For Each ws In Worksheets

'Variable declaration
Dim ticker As String
Dim opening_price As Double
Dim closing_price As Double
Dim percent_change As Integer
Dim total_stock_volume As Long
Dim starting_row As Integer
Dim yearly_change As Double
Dim percentage_change As Double

'Assigns new headers for the results
   ws.Cells(2, 9).Value = "Ticker"
   ws.Cells(2, 10).Value = "Yearly Change"
   ws.Cells(2, 11).Value = "Percent Change"
   ws.Cells(2, 12).Value = "Total Stock Volume"

   'Variable assignment
    total_stock_volume = 0
    starting_row = 2
    yearly_change = 0
    percentage_change = 0

   'Finding all the rows
    last_row = 1
    While ws.Cells(last_row, 1) <> ""
      last_row = last_row + 1
    Wend


    For i = 2 To last_row
       'Getting each ticker symbol
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          ticker = ws.Cells(i, 1).Value
          ws.Cells(3, 9).Value = ticker
          
          'Calculations
          opening_price = ws.Cells(i, 3).Value
          closing_price = ws.Cells(i, 6).Value
          yearly_change = yearly_change + (closing_price - opening_price)
          percentage_change = percentage_change + (yearly_change / opening_price * 100)
          total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
       
          'Print values
          ws.Range("I" & last_row).Value = ticker
          ws.Range("J" & last_row).Value = yearly_change
            If yearly_change < 0 Then
               ws.Range("J" & last_row).Interior.ColorIndex = 3
            Else
               ws.Range("J" & last_row).Interior.ColorIndex = 4
            End If
          ws.Range("K" & last_row).Value = initial_percentage_change
          ws.Range("L" & last_row).Value = initial_total_stock_volume
          
          'Reset variables
          last_row = last_row + 1
          yearly_change = 0
          percentage_change = 0
          total_stock_volume = 0
       
       Else
          yearly_change = yearly_change + (closing_price - opening_price)
          percentage_change = percentage_change + (yearly_change / opening_price * 100)
          total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
       
          'Print values
           ws.Range("I" & last_row).Value = ticker
           ws.Range("J" & last_row).Value = yearly_change
             If yearly_change < 0 Then
                ws.Range("J" & last_row).Interior.ColorIndex = 3
             Else
                ws.Range("J" & last_row).Interior.ColorIndex = 4
             End If
           ws.Range("K" & last_row).Value = percentage_change
           ws.Range("L" & last_row).Value = total_stock_volume
       End If
    
    Next i
Next ws

End Sub



