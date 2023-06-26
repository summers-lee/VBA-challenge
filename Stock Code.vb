Sub Stocks()
  ' Set the variables
  Dim Ticker As String
  Dim Yearly_Change As Double
  Dim Percentage_Change As Double
  Dim Total_Stock_Volume As Double
  Dim openPrice As Double
  Dim closePrice As Double
  Dim percentMin As Double
  Dim percentMax As Double
  Dim volumeMaxTicker As Double
  Dim Summary_Table_Row As Long
  Dim counter As Long

  
  ' Initialize variables before the loop
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  counter = 2
  Yearly_Change = 0
  Percentage_Change = 0
  Total_Stock_Volume = 0
  Summary_Table_Row = 2


' Loop through all sheets
   For Each ws In Worksheets
        ' Make Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.cells(2, 15).value = "Greatest % Increase"
        ws.cells(3, 15).value = "Greatest % Decrease"
        ws.cells(4, 15).value = "Greatest Total Volume"

    ' Loop through all the Ticker symbols
    For i = 2 To LastRow
    ' Check if we are still within the same Ticker, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value
      ' Calculate Yearly Change and save in column J. Also, highlight cell red (negative) or green (positive).
            closePrice = ws.cells(i, 6).value
            openPrice = ws.cells(i, 6).value
            Yearly_Change = closePrice - openPrice
            ws.cells(i, 10).value = Yearly_Change
            If Yearly_Change < 0 Then
                ws.cells(i, 10).Interior.ColorIndex = 3
                ws.cells(i, 11).Interior.ColorIndex = 3
            ElseIf Yearly_Change > 0 Then
                ws.cells(i, 10).Interior.ColorIndex = 4
                ws.cells(i, 11).Interior.ColorIndex = 4
            End If
            ' Calculate percent change and save in column K. Careful when dividing by zero!
            If Yearly_Change = 0 Or openPrice = 0 Then
                ws.cells(i, 11).value = 0
            Else
                ws.cells(i, 11).value = Format(Yearly_Change / openPrice, "#.##%")
            End If
      
      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
      ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      Total_Stock_Volume = 0

        ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  ' Find the values for greatest decrease/increase and greatest volume.
            If ws.cells(counter, 11).value > percentMax Then
                If ws.cells(counter, 11).value = ".%" Then
                Else
                    percentMax = ws.cells(counter, 11).value
                    percentMaxTicker = ws.cells(counter, 9).value
                End If
            ElseIf ws.cells(counter, 11).value < percentMin Then
                percentMin = ws.cells(counter, 11).value
                percentMinTicker = ws.cells(counter, 9).value
            ElseIf ws.cells(counter, 12).value > volumeMax Then
                volumeMax = ws.cells(counter, 12).value
                volumeMaxTicker = ws.cells(counter, 9).value
            End If
            ' Reset variables and go to next ticker symbol.
            counter = counter + 1
            summ = 0
            priceFlag = True

' Save the values for greatest decrease/increase and greatest volume.
    ws.cells(2, 17).value = Format(percentMax, "#.##%")
    ws.cells(3, 17).value = Format(percentMin, "#.##%")
    ws.cells(4, 17).value = volumeMax 
' Place corresponding ticker symbol to challance values.
    ws.cells(2, 16).value = percentMaxTicker
    ws.cells(3, 16).value = percentMinTicker
    ws.cells(4, 16).value = volumeMaxTicker

End Sub