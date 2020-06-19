Sub stock()

For Each ws In Worksheets
'Summary table 1
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"

'Summary table 2
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"


Dim last_row As Integer
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

total_vol = 0

Dim summary_row As Integer
summary_row = 2

Dim open_price As Double
Dim close_price As Double

open_price = 0
close_price = 0

open_price = ws.Cells(2, 3).Value
'MsgBox (open_price)


For i = 2 To lastRow

'if ticker names are different than the next cell and do the actions
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
   
       'set close_price on the current row
        close_price = ws.Cells(i, 6).Value
        
        'display ticker names
        ws.Range("J" & summary_row) = ws.Cells(i, 1).Value
        
        'calulate yearly change
        price_diff = (close_price - open_price)
        ws.Range("K" & summary_row) = price_diff
        
        'colorfill change percent
        If price_diff >= 0 Then
        ws.Range("K" & summary_row).Interior.ColorIndex = 4
        Else
        ws.Range("K" & summary_row).Interior.ColorIndex = 3
        End If
        
        
        'calculate change percent
        If open_price <> 0 Then
        ws.Range("L" & summary_row) = Format((price_diff / open_price), "Percent")
        Else
        'MsgBox ("there's a zero on open price")
        ws.Range("L" & summary_row) = Format(0, "percent")
        End If
        
        'calculate total volume
       total_vol = total_vol + ws.Cells(i, 7).Value
       ws.Range("M" & summary_row).Value = total_vol
       total_vol = 0
        
        'move on to the next row on the 1st summary table
        summary_row = summary_row + 1
        
        'reset values
        open_price = ws.Cells(i + 1, 3).Value
        close_price = 0
        price_diff = 0

   'else when the ticker names are the same, then     
   Else
   
        total_vol = total_vol + ws.Cells(i, 7).Value
         
           
   End If
Next i

'initialize the values for the 2nd summary table
max_increase_percent = ws.Range("L2").Value
max_decrease_percent = ws.Range("L2").Value
max_vol = ws.Range("M2").Value

'finding greatest % increase
For j = 2 To (summary_row - 1)
    If ws.Cells(j, 12) >= max_increase_percent Then
        max_increase_percent = ws.Cells(j, 12).Value
         ws.Range("Q2").Value = ws.Cells(j, 10).Value
    End If
Next j

'finding greatest % decrease
For x = 2 To (summary_row - 1)
    If ws.Cells(x, 12) < max_decrease_percent Then
        max_decrease_percent = ws.Cells(x, 12).Value
         ws.Range("Q3").Value = ws.Cells(x, 10).Value
    End If
Next x

'finding greate total volume
For y = 2 To (summary_row - 1)
   If ws.Cells(y, 13) > max_vol Then
      max_vol = ws.Cells(y, 13).Value
      ws.Range("Q4") = ws.Cells(y, 10).Value
    End If
Next y

'fill in the value into the 2nd summary table
ws.Range("R2") = Format(max_increase_percent, "percent")
ws.Range("R3") = Format(max_decrease_percent, "percent")
ws.Range("R4") = max_vol

'Auto fit width of all columns,easier to read
ws.Cells.EntireColumn.AutoFit

Next ws

End Sub




