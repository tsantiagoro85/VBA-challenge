Sub stock_data()

' Loop through each worksheet
For Each ws In Worksheets

  ' Set initial variables for ticker, quaterly change, percent change, total stock volume, close price and open price
    Dim ticker As String
    Dim quaterly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim close_price As Double
    Dim open_price As Double
    
  ' Set total_stock_volume to 0
    total_stock_volume = 0
  
  ' Specify open price column
    open_price = ws.Range("C2").Value

  ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  ' Create a second Summary Table that includes the ticker and values corresponding to greatest and lowest percent increase, and greatest total volume
    Dim Summary_Table_Row_2 As Integer
    Summary_Table_Row_2 = 2
  
  'Label the headers for the first summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quaterly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total_Stock_Volume"
  
  'Set last row so that the number of tickers does not have to be hard coded
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all the tickers
    For i = 2 To LastRow

    ' Check if we are still within the same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker
        ticker = ws.Cells(i, 1).Value
      
      ' Add the total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      
      ' Loop through the close price
        close_price = ws.Cells(i, 6).Value

      ' Calculate the quaterly change
        quaterly_change = (close_price - open_price)
      
      ' Print the ticker, total stock volume and quaterly change in the first summary table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
        ws.Range("J" & Summary_Table_Row).Value = quaterly_change
      
      ' Set the percentage change and specify that:
        If open_price = 0 Then
            
            percent_change = 0
        
        Else
      
      ' Add the percent change
        percent_change = (quaterly_change / open_price)

       End If

      ' Print the percent change in the Summary Table and set it up as percent
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock volume
      total_stock_volume = 0
      
      ' Reset the open price
        open_price = ws.Cells(i + 1, 3)

    Else

      ' Add to the percent change
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

    End If

  Next i
  
  ' Specify the conditional formating to color cells based on positive or negative quaterly change values
  ' Need to set up last row for Summary Table
   lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
        For i = 2 To lastrow_summary_table
        
            If ws.Cells(i, 10).Value > 0 Then
    
                ' Set the color to green
                  ws.Cells(i, 10).Interior.ColorIndex = 4
          
            Else
        
                ' Set the color to red
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
             End If
        
    Next i
 
   ' Label the Summary Table 2 headers
        ws.Range("O1").Value = "Category"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
    ' Label the Summary Table 2 categories
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Lowest % increase"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Identify the maximum and minimum values for the categories described above, and print in specified cells
         greatest_percent_increase = Application.WorksheetFunction.Max(ws.Range("K:K").Value)
         lowest_percent_increase = Application.WorksheetFunction.Min(ws.Range("K:K").Value)
         greatest_total_volume = Application.WorksheetFunction.Max(ws.Range("L:L").Value)
         
    ' Print values from above in specific cells within the second Summary Table
         ws.Range("Q2").Value = greatest_percent_increase
         ws.Range("Q3").Value = lowest_percent_increase
         ws.Range("Q4").Value = greatest_total_volume
         
   ' Create formulas to find the rows corresponding to the greatest and lowest percent increase, and greatest total volume
        Row_Max = Application.WorksheetFunction.Match(greatest_percent_increase, ws.Range("K:K").Value, 0)
        Row_Min = Application.WorksheetFunction.Match(lowest_percent_increase, ws.Range("K:K").Value, 0)
        Vol_Max = Application.WorksheetFunction.Match(greatest_total_volume, ws.Range("L:L").Value, 0)

    ' Find stock ticker using the information from above
        max_ticker = ws.Range("I" & Row_Max).Value
        min_ticker = ws.Range("I" & Row_Min).Value
        max_vol_ticker = ws.Range("I" & Vol_Max).Value

    ' Print tickers identified above in specific cells within the Summary Table
        ws.Cells(2, 16).Value = max_ticker
        ws.Cells(3, 16).Value = min_ticker
        ws.Cells(4, 16).Value = max_vol_ticker
        
    Next ws
        
End Sub
