Sub stock_analysis_vba_homework_MOSHER()

'To run the vba script on all worksheets in the excel file
For Each ws In Worksheets

    'tell vba the variable types
    'number of columns of raw data in the worksheet
    Dim ColumnCount As Long
    'The first ticker symbol in the spreadsheet
    Dim StockSymbol As String
    'The start for the comparisons of ticker symbols
    Dim SameSymbol As String
    'The first stock volume in the spreadsheet
    Dim StartVolume As Double
    'The start for the adding of stock volume
    Dim StockVolumeCounter As Double
    'The accumulation of adding up stock volumes
    Dim VolumeSum As Double
    'The opening stock price
    Dim OpeningPrice As Double
    'The closing stock price (always one column above the start of the next stock)
    Dim ClosingPrice As Double
    'The change in stock price
    Dim StockChange As Double
    'The percent change in stock price
    Dim PercentChange As Double

    'add the headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Opening Price"
    ws.Cells(1, 11).Value = "Closing Price"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"

    'starting values
    StockVolumeCounter = 0
    Results = 1
    ColumnCount = 1

    'set the starting ticker symbol
    SameSymbol = ws.Range("A2").Value
    'set the opening stock price
    OpeningPrice = ws.Range("C2").Value

    'find the last row of data
    finalrow = ws.Cells(Rows.Count, ColumnCount).End(xlUp).Row

    'create a list of stock symbols and a sum of stock volumes
    'the final row plus 1 to include the final stock in the worksheet
    For i = 2 To finalrow + 1
        
        StockSymbol = ws.Cells(i, 1).Value
        StartVolume = ws.Cells(i, 7).Value
        
        If StockSymbol = SameSymbol Then
            'Record the symbol name and add the stock volume to the previously added stock volume
            SameSymbol = ws.Cells(i, 1).Value
            VolumeSum = StockVolumeCounter + StartVolume
            StockVolumeCounter = VolumeSum
        Else
            'print ticker symbol and total accumulation of stock volume
            ws.Cells(Results + 1, 9).Value = SameSymbol
            ws.Cells(Results + 1, 14).Value = VolumeSum
            
            'calculate change in stock price and print values to worksheet
            'I printed opening price and closing price in case I want to use a pivot chart to better understand this data
            ClosingPrice = ws.Cells(i - 1, 6).Value
            StockChange = ClosingPrice - OpeningPrice
            ws.Cells(Results + 1, 10).Value = OpeningPrice
            ws.Cells(Results + 1, 11).Value = ClosingPrice
            ws.Cells(Results + 1, 12).Value = StockChange
            
            'calculate the percent change in stock price and print to worksheet
            'issue dividing by zero, printed the zero in 'opening price' if that was the case
            If OpeningPrice = 0 Then
                ws.Cells(Results + 1, 13).Value = 0
            Else
                PercentChange = StockChange / OpeningPrice
                ws.Cells(Results + 1, 13).Value = PercentChange
                ws.Cells(Results + 1, 13).NumberFormat = "0.00%"
            End If

            'Add green if stock change is greater than 0 and red if less than 0
            If StockChange >= 0 Then
                ws.Cells(Results + 1, 12).Interior.ColorIndex = 4
            Else
                ws.Cells(Results + 1, 12).Interior.ColorIndex = 3
            End If

        'Reset values
        StockVolumeCounter = 0
        SameSymbol = ws.Cells(i, 1).Value
        Results = Results + 1
        OpeningPrice = ws.Cells(i, 3).Value
        
        End If

    Next i

Next ws

End Sub

