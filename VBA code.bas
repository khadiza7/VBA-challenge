Attribute VB_Name = "Module1"
Sub StockMarketAnalysis():

'Variable declaration
Dim ws As Variant
Dim ticker As String
Dim starting_row As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim volume As Double

'Loop through each worksheet
For Each ws In Worksheets

  'Set opening price as the first price in the column
  opening_price = ws.Cells(2, 3).Value

  'Write the headers for the new table
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
 
   'Row of where the new data should be written
   starting_row = 2

   'Find the last row
   lastrow = 1
   
   While ws.Cells(lastrow, 1) <> ""
         lastrow = lastrow + 1
   Wend
     
     'Loop through each row in the column
     For i = 2 To lastrow
     
     'Add up the volume until a new ticker is found
     volume = volume + ws.Cells(i, 7).Value
     
         'If a new ticker is found print it to the new table and assign the closing price
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             ticker = ws.Cells(i, 1).Value
             closing_price = ws.Cells(i, 6).Value
             
             'Calculate yearly change and percentage change
             yearly_change = closing_price - opening_price
             
             'There are 0's in the column so ignore the 0's
             If opening_price <> 0 Then
                    percent_change = (yearly_change / opening_price) * 100
                
             End If
            
             
            ' Print the results in the new table
             ws.Cells(starting_row, 9).Value = ticker
             ws.Cells(starting_row, 10).Value = yearly_change
             ws.Cells(starting_row, 11).Value = percent_change
             ws.Cells(starting_row, 12).Value = volume
             
             'Formatting
             ws.Cells(starting_row, 11).Value = Format(ws.Cells(starting_row, 11) / 100, "#.####%")
             
             Dim cell As Variant
             cell = ws.Cells(starting_row, 10).Value
             Set cell = ws.Cells(starting_row, 10)
             
                If cell < 0 Then
                   cell.Interior.ColorIndex = 3
                ElseIf cell > 0 Then
                   cell.Interior.ColorIndex = 4
        
                End If
        
            
            ' Add one to the table row
             starting_row = starting_row + 1
             
             'Change the opening price for the new ticker
             opening_price = ws.Cells(i + 1, 3).Value
             
            volume = 0

          End If
      Next i
      

Next ws

End Sub
