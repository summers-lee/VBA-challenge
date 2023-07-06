Sub Stocks():

'Set all dimensions
Dim ws As Worksheet
Dim total As Double
Dim i As Long
Dim YearlyChange As Double
Dim j As Integer
Dim counter As Long
Dim LastRow As Long
Dim PerChange As Double
Dim days As Integer

Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestTotalTicker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double

   
   'Loop through each worksheet (Next ws noted already)
   For Each ws In Worksheets
      ws.Activate
  
   'Add column and summary box names
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
   Range("O2").Value = "Greatest % Increase"
   Range("O3").Value = "Greatest % Decrease"
   Range("O4").Value = "Greatest Total Volume"
   
   'Initial values and Last Row
   j = 0
   total = 0
   YearlyChange = 0
   counter = 2
   
   LastRow = Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To LastRow
   
   ' See if we are still within the same ticker sign
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           'Stores results
           total = total + Cells(i, 7).Value
           'Zero total volume
           If total = 0 Then
               'put results in the fields
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
            Else
              'Find non zero starting value
              If Cells(counter, 3) = 0 Then
               For find_value = counter To i
                       If Cells(find_value, 3).Value <> 0 Then
                           counter = find_value
                           Exit For
                       End If
                Next find_value
              End If
               
            'Calculate change
            YearlyChange = (Cells(i, 6) - Cells(counter, 3))
            PerChange = YearlyChange / Cells(counter, 3)
            'Start of the next stock ticker and record results
            counter = i + 1
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = YearlyChange
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = PerChange
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = total
            
                'Add color coding to YearlyChange
                If (YearlyChange > 0) Then
                Range("J" & 2 + j).Interior.ColorIndex = 4
                
                ElseIf (YearlyChange <= 0) Then
                Range("J" & 2 + j).Interior.ColorIndex = 3
                
                End If
                
                'Add color coding to Percent Change
                If (PerChange > 0) Then
                Range("K" & 2 + j).Interior.ColorIndex = 4
                
                ElseIf (PerChange <= 0) Then
                Range("K" & 2 + j).Interior.ColorIndex = 3
                
                End If
               
               
            End If
                         
           'Move on to next ticker symbol
           total = 0
           YearlyChange = 0
           j = j + 1
           days = 0
           'Add results for each ticker symbol
       Else
           total = total + Cells(i, 7).Value
    End If
    
          
    Next i
     
     
    'counter looking for greatest % increase
    LastRow = Cells(Rows.Count, "I").End(xlUp).Row
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotal = 0
    GreatestIncreaseTicker = ""
    GreatestDecreaseTicker = ""
    GreatestTotalTicker = ""
    
    For i = 2 To LastRow
    
    'want to compare the %Change to greatest increase value
    'want to do the same for greatest decrease
        If Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncreaseTicker = Cells(i, 9).Value
            GreatestIncrease = Cells(i, 11).Value
        End If
        
        If Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecreaseTicker = Cells(i, 9).Value
            GreatestDecrease = Cells(i, 11).Value
        End If
        
        If Cells(i, 12).Value > GreatestTotal Then
            GreatestTotalTicker = Cells(i, 9).Value
            GreatestTotal = Cells(i, 12).Value
        End If
         
    Next
    
    Cells(2, 16).Value = GreatestIncreaseTicker
    Cells(2, 17).Value = GreatestIncrease
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = GreatestDecreaseTicker
    Cells(3, 17).Value = GreatestDecrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = GreatestTotalTicker
    Cells(4, 17).Value = GreatestTotal
    Cells(4, 17).NumberFormat = "#"

   
    Next ws

End Sub