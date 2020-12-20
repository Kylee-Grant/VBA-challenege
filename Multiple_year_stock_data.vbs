Sub Market()
    
    'Create variable to hold worksheets
    Dim ws As Worksheet
    
    ' Set variables for parameters of the data
    Dim Last_Row As Double
    'Dim Last_Column As Double
    
    ' Initiate a summary table variable for tracking the row
    Dim Summary_Table_Row As Long
    
    ' Set an initial variable for holding the ticker
    Dim Ticker_Name As String
    
    ' Set an initial variable for holding the ticker volume, start price, end price, and yearly change and %
    Dim Stock_Volume As LongLong
    Dim Start_Price As Double
    Dim End_Price As Double
    Dim Yearly_Change As Double
    Dim Change_Percent As Double
    
    'Iterate through worksheets
    For Each ws In Worksheets
         
         ' Create column headers for the table
         ws.Cells(1, 8).Value = "Ticker"
         ws.Cells(1, 9).Value = "Yearly Change"
         ws.Cells(1, 10).Value = "Percent Change"
         ws.Cells(1, 11).Value = "Total Stock Volume"
         
         ' Keep track of the location for each stock in the summary table
          Summary_Table_Row = 2
          
         ' Set the location for the first start price of the worksheet
           Start_Price = ws.Cells(2, 3).Value
        
        ' Find the last row and column of the worksheet
         Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
        ' Loop through all stock values
         For i = 2 To Last_Row
    
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
              ' Set the ticker and print the ticker in the Summary Table
              Ticker_Name = ws.Cells(i, 1).Value
              ws.Range("H" & Summary_Table_Row).Value = Ticker_Name
        
              ' Add to the volume total and print the volume total to the Summary Table
               Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
               ws.Range("K" & Summary_Table_Row).Value = Stock_Volume
               
               'Add condition to avoid dividing by zero
               If Start_Price = 0 Then
                    ' Print the price difference to the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = 0
                    ' Print the change percent to the Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = 0
               Else
                    ' Grab end price
                     End_Price = ws.Cells(i, 6).Value
                    ' Calculate the percent change
                     Yearly_Change = (End_Price - Start_Price)
                     'Calculating percent. Note it is not multiplied by 100 due to formatting below.
                     Change_Percent = (Yearly_Change / Start_Price)
                     ' Print the price difference to the Summary Table
                     ws.Range("I" & Summary_Table_Row).Value = Yearly_Change
                     ' Print the change percent to the Summary Table
                     ws.Range("J" & Summary_Table_Row).Value = FormatPercent(Change_Percent, 2)
                     ' Initiate conditional formatting (green if > 0, red if < 0)
                     If Yearly_Change > 0 Then
                        ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf Yearly_Change < 0 Then
                        ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                        'If neither conditions met, the value was equal to 0. No coloring applied.
                End If
        
              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              'Set the starting price for the next ticker
              Start_Price = ws.Cells(i + 1, 3).Value
              
              ' Just in case, reset the volume total, yearly change, change percent, end price...
              ' Note, testing of the data did not indicate that these needed to be cleared for each ws
               Stock_Volume = 0
               Yearly_Change = 0
               Change_Percent = 0
               End_Price = 0
        
            ' If the cell immediately following a row is the same ticker...
            Else
        
              ' Add to the volume total
               Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
            End If
    
      Next i
  
  Next ws

End Sub
