Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

Instructions
Create a script that loops through all the stocks for one year and outputs the following information:
The ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
The total stock volume of the stock. The result should match the following image:# VBA-challenge

Here is my VBA Code:
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
  Dim priceFlag As Boolean
  
  
  ' Initialize variables before the loop
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  counter = 2
  Yearly_Change = 0
  Percentage_Change = 0
  Total_Stock_Volume = 0
  Summary_Table_Row = 2
  priceFlag = True
  
        ' Make Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"

    ' Loop through all the Ticker symbols
    For i = 2 To LastRow
    ' Check if we are still within the same Ticker, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the Ticker
      Ticker = Cells(i, 1).Value
      ' Calculate Yearly Change and save in column J. Also, highlight cell red (negative) or green (positive).
            closePrice = Cells(i, 6).Value
            openPrice = Cells(i, 3).Value
            Yearly_Change = closePrice - openPrice
            Cells(i, 10).Value = Yearly_Change
            If Yearly_Change < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
                Cells(i, 11).Interior.ColorIndex = 3
            ElseIf Yearly_Change > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
                Cells(i, 11).Interior.ColorIndex = 4
            End If
            ' Calculate percent change and save in column K. Careful when dividing by zero!
            If Yearly_Change = 0 Or openPrice = 0 Then
                Cells(i, 11).Value = 0
            Else
                Cells(i, 11).Value = Format(Yearly_Change / openPrice, "#.##%")
            End If
      
      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      Range("K" & Summary_Table_Row).Value = Percentage_Change
      ' Print the Total Stock Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      Total_Stock_Volume = 0

        ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i
  ' Find the values for greatest decrease/increase and greatest volume.
            If Cells(i, 11).Value > percentMax Then
                If Cells(i, 11).Value = ".%" Then
                Else
                    percentMax = Cells(i, 11).Value
                    percentMaxTicker = Cells(i, 9).Value
                End If
            ElseIf Cells(i, 11).Value < percentMin Then
                percentMin = Cells(i, 11).Value
                percentMinTicker = Cells(i, 9).Value
            ElseIf Cells(i, 12).Value > volumeMax Then
                volumeMax = Cells(i, 12).Value
                volumeMaxTicker = Cells(i, 9).Value
            End If
            ' Reset variables and go to next ticker symbol.
            counter = counter + 1
            summ = 0
            priceFlag = True

' Save the values for greatest decrease/increase and greatest volume.
    Cells(2, 17).Value = Format(percentMax, "#.##%")
    Cells(3, 17).Value = Format(percentMin, "#.##%")
    Cells(4, 17).Value = volumeMax
' Place corresponding ticker symbol to challance values.
    Cells(2, 16).Value = percentMaxTicker
    Cells(3, 16).Value = percentMinTicker
    Cells(4, 16).Value = volumeMaxTicker

End Sub
