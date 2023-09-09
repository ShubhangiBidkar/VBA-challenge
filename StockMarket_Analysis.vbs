
Sub stock()

'Define variable of the worksheet to excute the code in all work sheet at once in the workbook
  Dim ws As Worksheet

'Define a variable for Ticker
  Dim Ticker As String

'Define a variable for year open
  Dim opening As Double

'Define a variable for year close
  Dim closing As Double

'Define a variable for yearly change
 Dim year_change As Double

'Define a variable for total stock volume
  Dim totalstockvol As Double

'Define a variable for percent change
  Dim perc_change As Double


'Define a variable to set up a row to start
    Dim starting As Double

'Define a variable to set the row for open price
  Dim openPriceRow As Double

'Define variables to count the total number of rows in coulumn A and K
    Dim lastrow, lastrowTable As Double

'Define variables for greatest increase, greatest decrease and greatest total stock volume
  Dim maxValue, minValue, maxTotalStockvol As Double

    
    'Creates the loop to go through each worksheet in the workbook
    For Each ws In Worksheets
     
        'activate the worksheet
        ws.Activate
         
        'insert the headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 12).Value = "Total Volume"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        
        
     ' Loop through all stocks, 2 to last row
        lastrow = Rows(Rows.Count).End(xlUp).Row
       
      'assign starting integer
        starting = 2
        openPriceRow = 1
        totalstockvol = 0
      
      
      'set the last row
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
            For i = 2 To lastrow

               'it Ticker is not the same as the row before
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    'Get Ticker in order
                       Ticker = ws.Cells(i, 1).Value
                    
                    'get the i to count up by one each time
                      openPriceRow = openPriceRow + 1
                            
                    
                    'get the value from first day open and last day close
                        opening = ws.Cells(openPriceRow, 3).Value
                    closing = ws.Cells(i, 6).Value
                    
                    
                    'sum the total stock volume
                        For j = openPriceRow To i
                            totalstockvol = totalstockvol + ws.Cells(j, 7).Value
                        Next j
                    
                    'open the data is 0 to avoid division by zero
                        If opening = 0 Then
                            perc_change = closing
                        Else
                            year_change = closing - opening
                            perc_change = year_change / opening
                        End If

                    'print in sum table
                     ws.Cells(starting, 9).Value = Ticker
                     ws.Cells(starting, 10).Value = year_change
                     ws.Cells(starting, 11).Value = perc_change
                     ws.Cells(starting, 11).NumberFormat = "0.00%"
                     ws.Cells(starting, 12).Value = totalstockvol
                    
                    'go to the next row
                        starting = starting + 1
                    
                    'reset the values
                        totalstockvol = 0
                     year_change = 0
                     perc_change = 0
                    
                    'reset the count for the open price
                     openPriceRow = i
                End If
            Next i
            
        'Conditional formatting columns colors for yearly change

            jlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
                For j = 2 To jlastrow
                    
                 'if conditional formatting
                    If ws.Cells(j, 10) > 0 Then
                      ws.Cells(j, 10).Interior.ColorIndex = 4
                   Else
                     ws.Cells(j, 10).Interior.ColorIndex = 3
                   End If
                
              Next j
            
        ' Assign names for summary table 2
                
                Range("N2").Value = "Greatest % Increase"
                Range("N3").Value = "Greatest % Decrease"
                Range("N4").Value = "Greatest Total Volume"
                Range("O1").Value = "Ticker"
                Range("P1").Value = "Value"
            
       'set the initial values
         maxValue = 0
         minValue = 0
         maxTotalStockvol = 0
      
      
    
    lastrowTable = Cells(Rows.Count, 11).End(xlUp).Row
    
    'loop through column k and L to find the greatest Increase,greatest decrease and greatest total volumn
    For i = 2 To lastrowTable
    
        If ws.Cells(i, 11) > maxValue Then
                maxValue = ws.Cells(i, 11)
                maxTicker = ws.Cells(i, 9)
            Else
                maxValue = maxValue
            End If
            
            If ws.Cells(i, 11) < minValue Then
                minValue = ws.Cells(i, 11)
                minTicker = ws.Cells(i, 9)
            Else
                minValue = minValue
            End If
            
            If ws.Cells(i, 12) > maxTotalStockvol Then
                maxTotalStockvol = ws.Cells(i, 12)
                maxTotalVolumeTicker = ws.Cells(i, 9)
            Else
                maxTotalStockvol = maxTotalStockvol
            End If
        
        Next i

    'Set the values in the cells
        Range("O2").Value = maxTicker
        Range("P2").Value = maxValue
        Range("P2").NumberFormat = "0.00%"
        Range("O3").Value = minTicker
        Range("P3").Value = minValue
        Range("P3").NumberFormat = "0.00%"
        Range("O4").Value = maxTotalVolumeTicker
        Range("P4").Value = maxTotalStockvol
        
       
        Columns("I:Q").AutoFit

    Next ws

End Sub






