Sub MultipleYearStockData()

    Dim current As Worksheet
    For Each current In Worksheets
    
    'set initial variables
    Dim Ticker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    
    'keet track of the location of summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'loop through all tickers and the volumes
    Last_Row = current.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Last_Row
    
        'check if we are still within the same ticker
        If current.Cells(i + 1, 1).Value <> current.Cells(i, 1).Value Then
        
        'Set the ticker
        Ticker_Name = current.Cells(i, 1).Value
        
        'add to the total volume
        Total_Volume = Total_Volume + current.Cells(i, 7).Value
        
        'print the each ticker in the summary table
        current.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        'print the total volume to the summary table
        current.Range("J" & Summary_Table_Row).Value = Total_Volume
        
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset the total volume
        Total_Volume = 0
        
    'if the ticker is not the same
    Else
    
     ' Add to the Total Volume
      Total_Volume = Total_Volume + current.Cells(i, 7).Value

    End If

  Next i
  
Next
    
End Sub

