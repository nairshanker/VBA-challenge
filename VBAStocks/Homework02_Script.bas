Attribute VB_Name = "vba_challenge"
Sub vba_challenge()

Dim ws As Worksheet

For Each ws In Worksheets
    ws.Select

'HW using logic used in class

' Set an initial variable for holding the ticker symbol
  Dim Ticker_symbol As String
  Dim LastRow As Long
  ' Set variables for volume
  Dim Total_stock_volume As Double
  Dim table_row As Integer
  Dim open_price As Double
  Dim close_price As Double
  Dim yearly_change As Double
  Dim ticker_counter As Integer
  Dim percent_change As Double
  
  '### CHALLENGE ###
  
Dim Max_Per_TickerLabel As String
Dim Min_Per_TickerLabel As String
Dim Greatest_total_volume_TickerLabel As String

Dim Min_Baseline As Double
Dim Max_Baseline As Double

    '### END CHALLENGE VARIABLE DECLARATION ###
 

  Total_stock_volume = 0
  table_row = 2
  ticker_counter = 0
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
    'Set the header labels for the new columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
  
  ' Loop through all ticker rows
  For i = 2 To LastRow

   'Loop through rows in column 1 to find unique ticker symbols
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker_symbol = Cells(i, 1).Value
'      ' Add to the Total_stock_volume
      Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
      ' Print the ticker symbol in the Summary Table
      Range("I" & table_row).Value = Ticker_symbol
      ' Print the Total stock volume to the Summary Table
      Range("L" & table_row).Value = Total_stock_volume
      Range("J" & table_row).Value = yearly_change
     
      close_price = Cells(i, 6).Value
      open_price = Cells(i - ticker_counter, 3).Value
      yearly_change = close_price - open_price
    'Input yearly change value in  column I
        Range("J" & table_row).Value = yearly_change
        
        
        If open_price <> 0 Then
            percent_change = yearly_change / open_price
        Else
            percent_change = 0
        End If
        
        Range("K" & table_row).Value = percent_change
        
    If Range("J" & table_row).Value < 0 Then
        Range("J" & table_row).Interior.ColorIndex = 3
                        
    Else
        Range("J" & table_row).Interior.ColorIndex = 4
        
    End If
    
    'Reset the totals
        Total_stock_volume = 0
        ticker_counter = 0
    ' Add one to the table row
        table_row = table_row + 1
    
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
      ticker_counter = 1 + ticker_counter
    End If
    
         
Next i

Range(Cells(2, 11), Cells(LastRow, 11)).Select
                Selection.Style = "Percent"
                Selection.NumberFormat = "0.00%"

ws.Columns("A:Q").AutoFit

'### CHALLENGE ###

LastRow_challenge = Cells(Rows.Count, 11).End(xlUp).Row

   
Min_Baseline = 0
Max_Baseline = 0


Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

LastRow_challenge = Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow_challenge

    If Cells(i, 11).Value < Min_Baseline Then
        Min_Baseline = Cells(i, 11).Value
        Min_Per_TickerLabel = Cells(i, 9).Value
        Range("P3") = Min_Baseline
        Range("O3") = Min_Per_TickerLabel
    End If
    
    If Cells(i, 11).Value > Max_Baseline Then
        Max_Baseline = Cells(i, 11).Value
        Max_Per_TickerLabel = Cells(i, 9).Value
        Range("P2") = Max_Baseline
        Range("O2") = Max_Per_TickerLabel
    End If

    If Cells(i, 12).Value > Max_Baseline Then
        Max_Baseline = Cells(i, 12).Value
        Greatest_total_volume_TickerLabel = Cells(i, 9).Value
        Range("P4") = Max_Baseline
        Range("O4") = Greatest_total_volume_TickerLabel
    End If
Next i

Range(Cells(2, 16), Cells(3, 16)).Select
                Selection.Style = "Percent"
                Selection.NumberFormat = "0.00%"
                

Next ws

End Sub

