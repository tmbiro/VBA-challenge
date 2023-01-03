Attribute VB_Name = "YearlyStockChange"
'Copyright Tifani Biro, December 30, 2022
'Created for Penn Data Science Bootcamp
'Generate a summary table for stocks that shows each stock's:
'   1) Raw and percent difference in price at the beginning of a year vs. the end of a year
'   2) Total volume traded for the year
'Assumes different years are contained in different sheets and does not account for fiscal quarter differences

'Start subprocedure
Sub StockChange()
    
    'Create a variable for sheet, ticker names, the year's opening price, closing price, price change in the summary table
    Dim Sheet As Worksheet
    Dim Ticker As String
    Dim Summary_Row As Integer
    Dim Opening As Double
    Dim Closing As Double
    Dim ChangeRaw As Double
    Dim ChangePercentage As Double
    Dim Volume As Variant
    Dim Opening_Row As Variant
        
    'Loop through sheets
    For Each Sheet In Worksheets
        
        'Set counters for summary row, opening row, and volume
        Summary_Row = 2
        Opening_Row = 2
        Volume = 0
        
        'Add a column headers for ticker, year change, percent change, and volume of the summary table
        Sheet.Range("I1").Value = "Ticker"
        Sheet.Range("J1").Value = "Year Change"
        Sheet.Range("K1").Value = "Percent Change"
        Sheet.Range("L1").Value = "Volume"

        'Create a variable for last row
        LastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
       'Cycle through Rows
        For RowN = 2 To LastRow

           'Calculate stock volume and print it in the Summary Table
           Volume = Volume + Cells(RowN, 7).Value
         
            'Check if the row after the current one has the same ticker
            If Cells(RowN + 1, 1).Value <> Cells(RowN, 1).Value Then
            
            'If the next row is a different ticker

                'Set the ticker symbol & print it in the Summary Table
                Ticker = Cells(RowN, 1).Value
                Sheet.Range("I" & Summary_Row).Value = Ticker
                
                'Set opening price
                Opening = Cells(Opening_Row, 3).Value
                
                'Set close price
                Closing = Cells(RowN, 6).Value
                
                'Calculate raw price change, print it in the Summary Table, and color cells based on value
                ChangeRaw = Closing - Opening
                Sheet.Range("J" & Summary_Row).Value = ChangeRaw
                
                    'If raw price is less than 0 (i.e., negative in value)...
                    If Sheet.Range("J" & Summary_Row).Value < 0 Then
                              
                        'Color the cells red
                        Sheet.Range("J" & Summary_Row).Interior.ColorIndex = 3
    
                    'If the cells are not less than 0 (i.e., postive in value)
                    Else
                            
                        'Color the cells green
                        Sheet.Range("J" & Summary_Row).Interior.ColorIndex = 4
                    
                    'Otherwise, do nothing (e.g., if cell is empty or stock didn't change in value)
                    End If
                
                'Calculate percent change and print it in the Summary Table as a percentage
                ChangePercentage = (ChangeRaw / Opening)
                Sheet.Range("K" & Summary_Row).Value = ChangePercentage
                Sheet.Range("K" & Summary_Row).Value = FormatPercent(Sheet.Range("K" & Summary_Row).Value)
                
                'Print summed volume in Summary Table
                Sheet.Range("L" & Summary_Row).Value = Volume
                
                'Add one to the summary table row so that the next new ticker is printed in its own row
                Summary_Row = Summary_Row + 1
        
                'Set the Opening_Row for the new stock, which will be the row following the current one
                Opening_Row = RowN + 1
        
                'Set the Volume back to 0
                Volume = 0
                
            'Otherwise, don't add anything to the Summary Table yet
            End If
            
        'Move onto the next row of data
        Next RowN
        
        'Calculate which stocks showed the max percent increase, decrease, and volume using the data in the Summary Table
        'Add a column headers for ticker and value and row headers for value type to the Max Table
        Sheet.Range("O1").Value = "Ticker"
        Sheet.Range("P1").Value = "Value"
        Sheet.Range("N2").Value = "Greatest % Increase"
        Sheet.Range("N3").Value = "Greatest % Decrease"
        Sheet.Range("N4").Value = "Greatest Total Volume"
        
        'Set all values in the Max Table to 0 for baseline comparisons
        Sheet.Range("P2:P4").Value = 0
        
        'Update last row to reflect the last row in the Summary Table
        LastRow = Sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
       'Cycle through Summary Table rows
        For RowN = 2 To LastRow
        
            'If the percent value in the Summary Table is greater then the cell below it and the value listed under "Greatest % Increase" in the Max Table...
            If Cells(RowN, 11).Value > Cells(RowN + 1, 11).Value And Cells(RowN, 11).Value > Sheet.Range("P2").Value Then
            
                'Add this value to the Max Table as the "Greatest % Increase"
                Sheet.Range("P2").Value = Cells(RowN, 11).Value
                Sheet.Range("O2").Value = Cells(RowN, 9).Value
                
            'Otherwise, if it's less then the cell below it and the value listed in the "Greatest % Decrease" in the Max Table...
            ElseIf Cells(RowN, 11).Value < Cells(RowN + 1, 11).Value And Cells(RowN, 11).Value < Sheet.Range("P3").Value Then
                
                'Add this value to the Max Table as the "Greatest % Decrease"
                Sheet.Range("P3").Value = Cells(RowN, 11).Value
                Sheet.Range("O3").Value = Cells(RowN, 9).Value

            'Otherwise, do nothing with the percentage value in the present row
            End If
            
            'If the volume in the Summary Table is greater then the cell below it and the value listed as the "Greatest Total Volume" in the Max Table...
            If Cells(RowN, 12).Value > Cells(RowN + 1, 12).Value And Cells(RowN, 12).Value > Sheet.Range("P4").Value Then
            
                'Add this value to the Max Table as the "Greatest Total Volume"
                Sheet.Range("P4").Value = Cells(RowN, 12).Value
                Sheet.Range("O4").Value = Cells(RowN, 9).Value
                
            'Otherwise, do nothing with the volume value in the present row
            End If
            
        'Move onto next row of data
        Next RowN
    
    'Move onto next sheet
    Next Sheet
    
'End subprocedure
End Sub