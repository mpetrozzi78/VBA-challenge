Sub WallStreet()
    
    'LOOP through all sheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
     'Name the headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    ' SETTING VARIABLES
    
        'Set an initial variable for holding the ticker name
        Dim Ticker As String
    
        'Set the First row of a stock
        Dim BlockedRow As Double
        BlockedRow = 2
    
        'Set the Last row
        Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set the variable to hold the Yearly Change
        Dim YearlyChange As Double
    
        'Set the variable to hold the Percent Change
        Dim PercentChange As Double
    
        'Set an initial variable for holding the total volume per stock ticker
        Dim TotalVolume As Double
        TotalVolume = 0

        'Keep track of the location for each stock ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Setting Variable for greatest percentages and Total Volume
        Dim max_increase As Double
        max_increase = -100000
        Dim max_increase_ticker As String
        Dim max_decrase As Double
        max_decrease = 100000
        Dim max_decrease_ticker As String
        Dim max_totalvolume As Double
        max_totalvolume = 0
        Dim max_totalvolume_ticker As String
        
  ' Loop through all stock ticKers
  For i = 2 To LastRow

    ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'TOTAL VOLUME PROCESS
        
            'Set the Stock Ticker
            Ticker = ws.Cells(i, 1).Value

            'Add to the Stock Ticker Total
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'Print the Stock Ticker type in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            'Print the Total Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
            
    'Calculation of the Yearly Change
            YearlyChange = ws.Cells(i, 6) - ws.Cells(BlockedRow, 3)
            
            'Print the Yearly Change in the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            
            'CONDITIONAL FORMATTING of the Yearly Change
                If YearlyChange >= 0 Then
                
                'Green when the YearlyChange is positive
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                
                'Red when the YearlyChange is negative
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
    'Conditional to check if Cells(BlockedRow, 3) is equal to zero
    If ws.Cells(BlockedRow, 3) <> 0 Then
                
        'Calculation of the Percent Change
                PercentChange = YearlyChange / ws.Cells(BlockedRow, 3)
        
                'Print the Percent Change in the Summary Table
                ws.Range("k" & Summary_Table_Row).Value = PercentChange
                ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
                
       'Conditionals for greatest % increase
       If ws.Cells(Summary_Table_Row, 11) >= max_increase Then
       
       max_increase = ws.Cells(Summary_Table_Row, 11)
       max_increase_ticker = ws.Cells(Summary_Table_Row, 9)
       
       Else
       
       End If
       
        'Conditionals for greatest % decrease
       If ws.Cells(Summary_Table_Row, 11) <= max_decrease Then
       
       max_decrease = ws.Cells(Summary_Table_Row, 11)
       max_decrease_ticker = ws.Cells(Summary_Table_Row, 9)
       
       Else
       
       End If
        
        Else
        
                'What happens if =0
                ws.Range("k" & Summary_Table_Row).Value = "Error"
        
       End If
       
       
         'Conditionals for greatest total volume
       If ws.Cells(Summary_Table_Row, 12) >= max_totalvolume Then
       
       max_totalvolume = ws.Cells(Summary_Table_Row, 12)
       max_totalvolume_ticker = ws.Cells(Summary_Table_Row, 9)
       
       Else
       
       End If
       
       
        'Add 1 to BlockedRow
        BlockedRow = i + 1

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      TotalVolume = 0
      
     

    ' If the cell immediately following a row is the same stock ticker...
    Else

      ' Add to the Total Volume
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value

    End If

  Next i
  
  'Print the % increases, decreases and total volume
  
  ws.Range("O2").Value = max_increase_ticker
  ws.Range("P2").Value = max_increase
  ws.Range("P2").NumberFormat = "0.00%"
  ws.Range("O3").Value = max_decrease_ticker
  ws.Range("P3").Value = max_decrease
  ws.Range("P3").NumberFormat = "0.00%"
  ws.Range("O4").Value = max_totalvolume_ticker
  ws.Range("P4").Value = max_totalvolume
  
  
  Next ws

End Sub



