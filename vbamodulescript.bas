Attribute VB_Name = "vbamodulescript"
'start with one sheet


'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.


Sub tickerloop():

'challenge to loop through all the sheets, add ws to all cell + range

    
For Each ws In Worksheets


'Set the variables for each

Dim tickername As String

Dim tickervolume As Double
tickervolume = 0

Dim summary_ticker_row As Integer
summary_ticker_row = 2

Dim open_price As Double
open_price = ws.Cells(2, 3).Value

Dim close_price As Double

Dim yearly_change As Double

Dim percent_change As Double


'Now do a summary table with headers starting in random column

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly change"
ws.Cells(1, 12).Value = "Percent change"
ws.Cells(1, 13).Value = "Total stock volume"


'count number of rows

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Now for the looping by ticker names

For i = 2 To lastrow


'next different cell

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

tickername = ws.Cells(i, 1).Value

tickervolume = tickervolume + ws.Cells(i, 7).Value

'now closing price and yearly change

   close_price = ws.Cells(i, 6).Value

    yearly_change = (close_price - open_price)

'print the above for summary table

ws.Range("J" & summary_ticker_row).Value = tickername

ws.Range("M" & summary_ticker_row).Value = tickervolume

ws.Range("K" & summary_ticker_row).Value = yearly_change


'Calculate percent change

If (open_price = 0) Then

          percent_change = 0
        

Else

       percent_change = yearly_change / open_price

End If

ws.Range("L" & summary_ticker_row).Value = percent_change
ws.Range("L" & summary_ticker_row).NumberFormat = "0.00%"

'reset to next row

summary_ticker_row = summary_ticker_row + 1

tickervolume = 0

open_price = ws.Cells(i + 1, 3)

Else

   tickervolume = tickervolume + ws.Cells(i, 7).Value

End If

Next i

'Conditional formatting colour- positive change green, negative red
    'last row of summary table
    
    lastrow_summary_table = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 11).Value > 0 Then
                
                ws.Cells(i, 11).Interior.ColorIndex = 4
            
            Else
                
                ws.Cells(i, 11).Interior.ColorIndex = 3
            
            End If
            
Next i

'BONUS


'Add functionality to your script to return the stock with the
'"Greatest % increase"
'"Greatest % decrease"
'"Greatest total volume"
'First label them


  ws.Cells(2, 16).Value = "Greatest % Increase"
  ws.Cells(3, 16).Value = "Greatest % Decrease"
  ws.Cells(4, 16).Value = "Greatest Total Volume"
  ws.Cells(1, 17).Value = "Ticker"
  ws.Cells(1, 18).Value = "Value"

For i = 2 To lastrow_summary_table

'calculate maximum percent change

 If ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
    ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
    ws.Cells(2, 18).Value = ws.Cells(i, 12).Value
    ws.Cells(2, 18).NumberFormat = "0.00%"

'exact same for minimum percent change using min fucntion

 ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & lastrow_summary_table)) Then
    ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
    ws.Cells(3, 18).Value = ws.Cells(i, 12).Value
    ws.Cells(3, 18).NumberFormat = "0.00%"

'now maximum total volume

   ElseIf ws.Cells(i, 13).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & lastrow_summary_table)) Then
          ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
          ws.Cells(4, 18).Value = ws.Cells(i, 13).Value

   End If

Next i

Next ws

End Sub
