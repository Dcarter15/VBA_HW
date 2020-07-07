Attribute VB_Name = "Module1"
Sub credit_card()

For Each ws In Worksheets

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Set an initial variable for holding the stock price
  Dim Stock_Total As Double
  Stock_Total = 0

  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all tickers
  For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      Ticker_Name = ws.Cells(i, 1).Value
      
      ' Calculate change in stock
        If ws.Cells(i, 2).Value = "20140101" And ws.Cells(i, 2).Value = "20141231" Then
            Stock_Total = ((Stock_Total + ws.Cells(i, 3).Value) - ws.Cells(i, 6).Value)
        End If
        
      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ' Print the change in stock to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock price
      Stock_Total = 0
      
      ' If the cell immediately following a row is the same ticker...
    Else

        If ws.Cells(i, 2).Value = "20140101" And ws.Cells(i, 2).Value = "20141231" Then
            Stock_Total = ((Stock_Total + ws.Cells(i, 3).Value) - ws.Cells(i, 6).Value)
        End If

    End If

  Next i

Next ws

End Sub

