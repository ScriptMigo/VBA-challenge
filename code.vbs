Sub multipleYearStockData()
    ' Purpose:
    ' * Create a script that will loop through all the stocks for one year for each run and take the following information.
    ' * The ticker symbol.
    ' * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' * The total stock volume of the stock.
    ' * You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    ' Variable Declarations
    Dim currentTickerSymbol         As String
    Dim yearStartOpeningValue       As Double
    Dim yearEndClosingValue         As Double
    Dim overallYearChange           As Variant
    Dim currentRow                  As Long
    Dim currentSummaryRow           As Integer
    Dim lastRow                     As Long
    Dim overallYearChangePercent    As Double
    Dim totalStockVolume            As Variant
    Dim ws                          As Worksheet
    
    ' Assign starting variable values
    
    currentTickerSymbol = ""
    yearStartOpeningValue = 0
    yearEndClosingValue = 0
    overallYearChange = 0
    currentRow = 2
    currentSummaryRow = 2
    overallYearChangePercent = 0
    totalStockVolume = 0
    lastRow = 0
    
    ' Loop sheets in document
    For Each ws In ThisWorkbook.Worksheets
        ' Activate worksheet
        ws.Select
        ws.Activate
        
        ' Define rows in sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Get initial opening value
        yearStartOpeningValue = Cells(2, 3)
        
        ' Loop rows in current sheet
        For currentRow = 2 To lastRow
            ' Get current ticker symbol
            currentTickerSymbol = Cells(currentRow, 1)
            
            ' Get our first ticker start price, this is only needed for the first iteration of each sheet
            'If currentRow = 2 Then
            '    yearStartOpeningValue = Cells(currentRow, 3)
            'End If
            
            ' Evaluate if next row is the same ticker symbol as the current
            If currentTickerSymbol = Cells(currentRow + 1, 1) Then
                ' Get values for current row and append to current tickers total counts
                totalStockVolume = Cells(currentRow, 7) + totalStockVolume
            Else
                ' This indicates the next ticker is not related to the current, grab the last closing value, the next ticker
                ' opening value, write the values, clear variables, assign new ticker symbol, increase current summary row
                yearEndClosingValue = Cells(currentRow, 6)
                totalStockVolume = Cells(currentRow, 7) + totalStockVolume
                
                ' Calculate overall year change
                overallYearChange = yearEndClosingValue - yearStartOpeningValue
                
                ' Calculate overall year percentage
                If yearStartOpeningValue > 0 Then
                    overallYearChangePercent = overallYearChange / yearStartOpeningValue * 100
                Else
                    overallYearChangePercent = 0
                End If
                
                ' set cell values
                ' ticker
                Cells(currentSummaryRow, 9).Value = currentTickerSymbol
                ' yearly change value and color
                Cells(currentSummaryRow, 10).Value = overallYearChange
                
                ' conditional formatting
                If overallYearChange > 0 Then
                    Cells(currentSummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(currentSummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' yearly change percentage
                Cells(currentSummaryRow, 11).Value = Str(Round(overallYearChangePercent, 2)) + "%"
                ' total stock volume
                Cells(currentSummaryRow, 12).Value = totalStockVolume
                ' clear variables
                overallYearChange = 0
                overallYearChangePercent = 0
                totalStockVolume = 0
                ' set new currentTickerSymbol
                currentTickerSymbol = Cells(currentRow + 1, 1)
                ' set new summary row
                currentSummaryRow = currentSummaryRow + 1
                
                ' Get new tickers starting value
                yearStartOpeningValue = Cells(currentRow + 1, 3)
            End If
        Next currentRow
        ' Clear summary row for next sheet
        currentSummaryRow = 2
        
        ' Now we loop over the values to find the min/max values for % change, as well as the greatest total volume
        MaxVolume = Application.WorksheetFunction.Max(Range("l:l"))
        MaxChange = Application.WorksheetFunction.Max((Range("k:k").Value)) * 100
        MinChange = Application.WorksheetFunction.Min((Range("k:k").Value)) * 100
        
        
        
        Cells(2, 15).Value = "MaxVolume"
        Cells(3, 15).Value = "MaxChange"
        Cells(4, 15).Value = "MinChange"
        Cells(2, 17).Value = MaxVolume
        Cells(3, 17).Value = Str(MaxChange) + "%"
        Cells(4, 17).Value = Str(MinChange) + "%"
        
    Next ws
End Sub

