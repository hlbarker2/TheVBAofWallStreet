Attribute VB_Name = "Module1"
Sub StockData()

'Loop through all sheets
Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

'Set variables
Dim ticker As String

Dim stock_volume As Double
stock_volume = 0

Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Keep track of location for each stock ticker in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Determine last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set initial open price
open_price = Cells(2, 3).Value

'Loop through all stock tickers
For i = 2 To lastrow

    'Check if we are still within the same stock ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Set the ticker name
    ticker = Cells(i, 1).Value
    
    'Set close price
    close_price = Cells(i, 6).Value
    
    'Calculate yearly change
    yearly_change = close_price - open_price
    
    'Calculate percent change
    If (open_price = 0 And close_price = 0) Then
    percent_change = 0
    ElseIf (open_price = 0 And close_price <> 0) Then
    percent_change = 1
    Else
    percent_change = yearly_change / open_price
    End If
    
    'Add to the stock total
    stock_volume = stock_volume + Cells(i, 7).Value
    
    'Print in the summary table
    Range("I" & summary_table_row).Value = ticker
    Range("J" & summary_table_row).Value = yearly_change
    
    Range("K" & summary_table_row).NumberFormat = "0.00%"
    Range("K" & summary_table_row).Value = percent_change
    
    Range("L" & summary_table_row).Value = stock_volume
    
    'Add one to the summary table row
    summary_table_row = summary_table_row + 1
    
    'Reset open price
    open_price = Cells(i + 1, 3).Value
    
    'Reset stock total
    stock_volume = 0
    
    'If the cell immediately following a row is the same stock
    Else
    
    'Add to the stock total
    stock_volume = stock_volume + Cells(i, 7).Value
    
    End If
    
Next i

'Determine last row for yearly change
yclastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Conditional formatting for yearly change
For j = 2 To yclastrow
    If Cells(j, 10).Value > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    Else
    Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

'Create summary table for greatest change and volume
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Determine greatest change and volume
For k = 2 To yclastrow
    If Cells(k, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & yclastrow)) Then
    Range("O2").Value = Cells(k, 9).Value
    Range("P2").Value = Cells(k, 11).Value
    Range("P2").NumberFormat = "0.00%"
    
    ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & yclastrow)) Then
    Range("O3").Value = Cells(k, 9).Value
    Range("P3").Value = Cells(k, 11).Value
    Range("P3").NumberFormat = "0.00%"
    
    ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & yclastrow)) Then
    Range("O4").Value = Cells(k, 9).Value
    Range("P4").Value = Cells(k, 12).Value
    
    End If

Next k
    
Next ws
    
End Sub
