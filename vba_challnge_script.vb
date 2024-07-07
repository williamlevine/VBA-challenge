Sub tickerCatch():

'set up worksheet variable and looping

Dim ws As Worksheet

For Each ws In Worksheets

'set up variables, initialize variables

Dim tickerName As String

Dim tickerVolTotal As Double
tickerVolTotal = 0

Dim tickerOpen As Double
Dim tickerClose As Double

Dim percentChange As Double

Dim i As Long

Dim summaryTableRow As Long
summaryTableRow = 2

'automatically find the last row for each worksheet

Dim lastrow As Long
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'set up column headers for each worksheet

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'loop through all rows and determine each row's location

    For i = 2 To lastrow
        
'if the row is the first of its ticker, set the open value and begin counting the volume total

        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

            tickerOpen = ws.Cells(i, 3).Value
            
            tickerVolTotal = ws.Cells(i, 7).Value

'if the row is the final of its ticker, then set the ticker name and add it to the summary table

        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            tickerName = ws.Cells(i, 1).Value
        
            ws.Cells(summaryTableRow, 9).Value = tickerName

'add its volume to the volume total, set the close value
        
            tickerVolTotal = tickerVolTotal + ws.Cells(i, 7).Value
        
            tickerClose = ws.Cells(i, 6).Value

'establish quarterlyChange as a variable, calculate it as the difference between the open and close, and add to summary table

            Dim quarterlyChange As Double
            quarterlyChange = tickerClose - tickerOpen
            ws.Cells(summaryTableRow, 10).Value = quarterlyChange

'determine whether quarterlyChange is positive or negative, and apply conditional formatting (red or green) accordingly
            
            If quarterlyChange > 0 Then
                
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(102, 255, 102)

            ElseIf quarterlyChange < 0 Then
            
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 64, 64)
                
            End If

'calculate percentChange as the difference between close and open divided by the open, format as a percent, and add to summary table

            percentChange = (tickerClose - tickerOpen) / tickerOpen
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"

'insert the volume total into the summary table

            ws.Cells(summaryTableRow, 12).Value = tickerVolTotal

'add one to the summary table row counter; reset the volume total to zero

            summaryTableRow = summaryTableRow + 1
            
            tickerVolTotal = 0

'if the row is neither the first nor last of its ticker, then add to the volume total sum

        Else
        
            tickerVolTotal = tickerVolTotal + ws.Cells(i, 7).Value
            
        End If
    
    Next i
    
'establish and initialize variables for the statistic summary table
    
    Dim j As Integer
    Dim lastrowSummary As Integer
    lastrowSummary = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    Dim greatestIncrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolume As Double
    Dim greatestVolumeTicker As String
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

'set up headers for statistic summary table

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

'begin looping through the summary table

      For j = 2 To lastrowSummary

'if the ticker has the greatest increase value yet encountered, then set it as the greatestIncrease and greatestIncreaseTicker

        If ws.Cells(j, 11).Value > greatestIncrease Then
        
        greatestIncrease = ws.Cells(j, 11).Value
        greatestIncreaseTicker = ws.Cells(j, 9).Value

        End If

'if the ticker has the greatest decrease value yet encountered, then set it as the greatestDecrease and greatestDecreaseTicker

        If ws.Cells(j, 11).Value < greatestDecrease Then
        
        greatestDecrease = ws.Cells(j, 11).Value
        greatestDecreaseTicker = ws.Cells(j, 9).Value
        
        End If

'if the ticker has the greatest volume yet encountered, then set it as the greatestVolume and greatestVolumeTicker

        If ws.Cells(j, 12).Value > greatestVolume Then
        
        greatestVolume = ws.Cells(j, 12).Value
        greatestVolumeTicker = ws.Cells(j, 9).Value
        
        End If
      
      Next j

'once the entire summary table has been combed for the greatest values, insert them into the summary stats table, format percentages where necessary        
        
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = greatestVolumeTicker
    ws.Cells(4, 17).Value = greatestVolume
        
'go to the next worksheet and repeat

  Next ws

End Sub