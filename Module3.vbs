Attribute VB_Name = "Module3"
Sub Ticker_Market()



For Each ws In Worksheets
'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest%Increase"
ws.Range("O3").Value = "Gretest%Decrease"
ws.Range("O4").Value = "Greatest Total Value"

'Define Ticker variable
Dim Tickername As String

'Set a variable to hold the total volume of ticker
Dim tickerVolume As Double
tickerVolume = 0

'Set Value to 0 for Gretest%
ws.Range("Q2").Value = 0
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = 0
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = 0
ws.Range("Q4").NumberFormat = "#,###,###,###,###"

'Set new variable for prices and percent changes
Dim open_price As Double
open_price = ws.Cells(2, 3).Value

Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

Dim summary_ticker_row As Integer
summary_ticker_row = 2

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Loop trough the rows by the ticker names
'Set initial and last row for worksheet

Dim i As Long
Dim j As Integer
For i = 2 To Lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Set the ticker name
        Tickername = ws.Cells(i, 1).Value
        tickerVolume = tickerVolume + ws.Cells(i, 7).Value
        'Print the ticker name in the summary table
        ws.Range("I" & summary_ticker_row).Value = Tickername
        'Print the trade volume for each ticker in the summary table
        Total_Stock_volume = tickerVolume
        ws.Range("L" & summary_ticker_row).Value = Total_Stock_volume
        'Collect information about closing price
        close_price = ws.Cells(i, 6).Value
        'Calculate yearly change
        yearly_change = (close_price - open_price)
        'Print the yearly change for each ticker in the summary table
        ws.Range("J" & summary_ticker_row).Value = yearly_change
        
        'Check for the non-divisibility condition when calculating the percent change
            If (open_price = 0) Then
                percent_change = 0
            Else
            
                percent_change = yearly_change / open_price
            
            End If
    
        'Print the yearly change for each ticker in the summary table
        ws.Range("K" & summary_ticker_row).Value = percent_change
        ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
        
        'Set % Increase or Decrease
        If percent_change > ws.Range("Q2").Value Then
            ws.Range("P2").Value = Tickername
            ws.Range("Q2").Value = percent_change
        End If
        
        If percent_change < ws.Range("Q3").Value Then
            ws.Range("P3").Value = Tickername
            ws.Range("Q3").Value = percent_change
        
        End If
        
        If tickerVolume > ws.Range("Q4").Value Then
            ws.Range("P4").Value = Tickername
            ws.Range("Q4").Value = tickerVolume
        End If
        
        
        'Reset the row counter.Add one to the summary_ticker_row
        summary_ticker_row = summary_ticker_row + 1
        'Reset volume of trade to zero
        tickerVolume = 0
        'Reset the opening price
        open_price = ws.Cells(i + 1, 3)

    Else

        'Add the volume of trade
        tickerVolume = tickerVolume + ws.Cells(i, 7).Value
    
    End If

Next i
'Conditional farmatting the will highlight positive change in green and negative chnge in red
'Find the last row of the summary table
lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Color code yearly change
For i = 2 To lastrow_summary_table
 If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 10
    
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
 Next i
Next
  
End Sub
