Sub StockData()
For Each ws In Worksheets

  ' Set an initial variable for holding the Ticker symbol
  Dim Ticker_Name As String

  ' Set an initial variable for holding the Volume per Ticker symbol
  Dim Ticker_Total_Volume As Double
  Ticker_Total_Volume = 0
  
  ' Set an initial variable for holding the yearly change per Ticker symbol
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  ' Set an initial variable for holding the percent change per Ticker symbol
  Dim Percent_Change As Double
  Percent_Change = 0
   
  ' Set an initial variable for holding the opening value
  Dim Opening_Value As Double
  Opening_Value = 0
  
  ' Set an initial variable for holding the last trade value
  Dim Closing_Value As Double
  Closing_Value = 0
    
  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Setting the first opening value location
  Dim j As Long
  j = 2
   
    
  ' Finds last entry or row in worksheet, so ws.range of rows is not hardcoded.
  ' Counts the number of rows
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Headers for new columns for each ws
    ws.Range("I" & 1).Value = "Ticker"
    ws.Range("J" & 1).Value = "YearlyChange"
    ws.Range("K" & 1).Value = "PercentChange"
    ws.Range("L" & 1).Value = "TotalStockVolume"
    
  'Later iteration LastRow will be end point of data in sheet
  'So the for loop below will be For i = 2 To lastRow which we established in variable above

  ' Loop through all rows of trades
  For i = 2 To LastRow
    
      ' Check if we are still within the same Ticker symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker symbol name
      Ticker_Name = ws.Cells(i, 1).Value
      
      'We need to get the opening value associated with the first loop - this will be the earliest date for the symbol
      Opening_Value = ws.Cells(j, 3).Value
           
      'Set closing value - we need to get the value associated with the last loop - this will be the closing price
      Closing_Value = ws.Cells(i, 6).Value
      
      ' Add to the Total Ticker volume
      Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
      
      ' Add to the yearly price change per ticker symbol
      Yearly_Change = Closing_Value - Opening_Value
    
      If Opening_Value = 0 Then
      Percent_Change = 0
      Else
      
      Percent_Change = (Yearly_Change / Opening_Value)
      End If
        
      ' Print the Ticker symbol name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ' Print the Ticker volume amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume
      
      'Tested these values below in order to work through syntax of my for loop
      'Print the Opening Value to the summary table
      'ws.Range("M" & Summary_Table_Row).Value = Opening_Value
      
      'Print the Closing Value to the summary table
      'ws.Range("N" & Summary_Table_Row).Value = Closing_Value
      
     ' Print the yearly price change amount to the Summary Table and set cell
     ' color based on positive (green) and negative (red) change in value.
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        If Yearly_Change > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
      ' Print the percent change amount to the Summary Table and set cell
      ' color based on positive (green) and negative (red) change in value.
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      If Percent_Change > 0 Then
        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker total volume
      Ticker_Total_Volume = 0
      
     'This is for the opening value
      j = i + 1

    ' If the cell immediately following a row is the same Ticker symbol...
    Else

      ' Add to the Ticker volume total
      Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                
     
    End If
        
  Next i
  
  'This is the bonus section where we want to find the greatest %increase, decrease and volume for each sheet
  
  'this section is for greatest volume
  Dim Max_Vol As Double
  Max_Vol = 0
  Dim Ticker_Sym As String
  
  For x = 2 To Summary_Table_Row
    If ws.Cells(x, 12).Value > Max_Vol Then
      Max_Vol = ws.Cells(x, 12).Value
      Ticker_Sym = ws.Cells(x, 9).Value
    End If
     ws.Range("N4").Value = "Max Total Volume"
     ws.Range("O4").Value = Ticker_Sym
     ws.Range("P4").Value = Max_Vol
  
  Next x
  
 'this section is for greatest % increase
 
 Dim GreatestIncrease As Double
  GreatestIncrease = 0
  
  For Z = 2 To Summary_Table_Row
    If ws.Cells(Z, 11).Value > GreatestIncrease Then
      GreatestIncrease = ws.Cells(Z, 11).Value
      Ticker_Sym = ws.Cells(Z, 9).Value
    End If
     ws.Range("N2").Value = "Greatest % increase"
     ws.Range("O2").Value = Ticker_Sym
     ws.Range("P2").Value = GreatestIncrease
  
  Next Z
  
 'this section is for greatest% decrease
 
  Dim GreatestDecrease As Double
  GreatestDecrease = 0
  
  For Q = 2 To Summary_Table_Row
    If ws.Cells(Q, 11).Value < GreatestDecrease Then
      GreatestDecrease = ws.Cells(Q, 11).Value
      Ticker_Sym = ws.Cells(Q, 9).Value
    End If
     ws.Range("N3").Value = "Greatest % decrease"
     ws.Range("O3").Value = Ticker_Sym
     ws.Range("P3").Value = GreatestDecrease
  
  Next Q
  
  Next ws

End Sub


