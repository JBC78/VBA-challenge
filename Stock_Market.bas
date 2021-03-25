Attribute VB_Name = "Module1"
Sub UpdateAllSheets()
'Update all worksheets
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
      ws.Select
       Call Wall_Street_VBA
   Next
    Application.ScreenUpdating = True
End Sub

Sub Wall_Street_VBA()
    
  
    Dim total As Double                                     'set dimensions
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim Opening As Long
    Dim Closing As Long
    Dim lastRow As Long
    Dim percentChange As Double
    Dim greatest_percent_increase As Double
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease As Double
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_total_volume As Long
    Dim greatest_total_volume_ticker As String
    Dim summaryRow As Integer
    summaryRow = 2
        
    Range("I1").Value = "Ticker"                            'set summary title row for Ticker
    Range("J1").Value = "Yearly Change"                     'set summary title row for Yearly Change
    Range("K1").Value = "Percent Change"                    'set summary title row for Percent Change
    Range("L1").Value = "Total Stock Volume"                'set summary title row for Total Stock Volume

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row          'get the row number of the last row with data
    For i = 2 To lastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then  'if ticker changes then print results
        
        ticker = Cells(i, 1).Value                          'store result as variable
        
        Range("I" & summaryRow).Value = ticker              'Add Ticker values
        open_value = Cells(summaryRow, 3).Value             'loop through column 3 and store open values
        closing_value = Cells(i, 6).Value                   'loop through column 6 and store close values
        Yearly_change = closing_value - open_value          'define yearly range by stored close and open values
        percent_change = Round((Yearly_change) / open_value, 4) * 100 'define percent change by yearly change / open value
                total_volume = Cells(i, 7).Value
        Cells(summaryRow, 10).Value = Yearly_change
        Cells(summaryRow, 11).Value = percent_change
        Cells(summaryRow, 12).Value = total_volume
        Summary = i + 1
               
         If Yearly_change > 0 Then
            Cells(summaryRow, 10).Interior.ColorIndex = 4 'Green for Positive Change in Stock
            ElseIf Yearly_change < 0 Then
            Cells(summaryRow, 10).Interior.ColorIndex = 3 'Red for Negitive Change in Stock
            End If
                    
        summaryRow = summaryRow + 1                         'Add one to summary row to loop to next row at end of conditions
               
    End If
    
Next i
   
    Range("P1").Value = "Ticker"                             'set Challenges titles Ticker, Value, Greatest % Increase/Decrease & Total Volume
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
    
    For i = 2 To lastRow                                    'Loop through the tickers in column "I" & find last row
    
    If Cells(i, 11).Value > greatest_percent_increase Then  'Find the ticker with the greatest percent increase
        greatest_percent_increase = Cells(i, 11).Value      'set greatest percent increase value
        greatest_percent_increase_ticker = Cells(i, 9).Value 'set greatest percent increase ticker value
        
        End If
                            
    If Cells(i, 11).Value < greatest_percent_decrease Then  'Find the ticker with the greatest percent decrease
        greatest_percent_decrease = Cells(i, 11).Value      'set greatest decrease percent value
        greatest_percent_decrease_ticker = Cells(i, 9).Value 'set greatest decrease ticker value
        End If
        
    If Cells(i, 12).Value > greatest_total_volume Then      'Find the ticker with the greatest total volume
        greatest_total_volume = Cells(i, 12).Value          'set greatest total volume value
        greatest_total_volume_ticker = Cells(i, 9).Value    'se greatest total volume ticker
        
        End If
    
Next i
    
    Range("P2").Value = greatest_percent_increase_ticker     'Add values for greatest percent increase ticker
    Range("Q2").Value = greatest_percent_increase             'Add values for greatest percent increase
    Range("P3").Value = greatest_percent_decrease_ticker     'Add values for greatest percent decrease ticker
    Range("Q3").Value = greatest_percent_decrease            'Add values for greatest percent decrease
    Range("P4").Value = greatest_total_volume_ticker         'Add values for greatest total volume ticker
    Range("Q4").Value = greatest_total_volume                'Add values for total volume
    
End Sub

