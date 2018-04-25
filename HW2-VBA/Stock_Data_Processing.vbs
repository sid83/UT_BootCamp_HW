Attribute VB_Name = "Module1"
Sub stockdata_processing()
For Each ws In Worksheets
    worksheetname = ws.Name
    Worksheets(worksheetname).Activate
    Range("i:r").Delete
    Dim linecount, i As Long
    Dim ii As Integer
    Dim stockstartline(4000), stockendline(4000) As Long
    Dim TickerSymbol(4000) As String            'initialized as dynamic array
    Dim OpeningPrice_year(4000), ClosingPrice_year(4000) As Double
    Dim YearlyChange(4000), PercentChange(4000) As Double
    Dim TotalVolume(4000) As LongLong
    Dim rownumber(4000) As Long
    'getting output table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    'counting number of columns
    linecount = Application.WorksheetFunction.CountA(Range("a:a"))
    'stockstartline = 2 'counter for picking price at beginning of the year
    stock_count = 0
    'ReDim stockstartline(1)
    stockstartline(0) = 2
    'loop for calculating row numbers of beginning and end of each stock
    For i = 2 To linecount - 1
        If Cells(i, 1) <> Cells(i + 1, 1) Then
        stock_count = stock_count + 1
        stockstartline(stock_count) = i + 1
        stockendline(stock_count - 1) = i
        End If
    Next i
    stockendline(stock_count) = linecount
    'loop to pick details of each stock
    For ii = 0 To stock_count - 1
        TickerSymbol(ii) = Cells(stockstartline(ii), 1).Value
        OpeningPrice_year(ii) = Cells(stockstartline(ii), 3).Value
        ClosingPrice_year(ii) = Cells(stockendline(ii), 6).Value
        YearlyChange(ii) = ClosingPrice_year(ii) - OpeningPrice_year(ii)
        If OpeningPrice_year(ii) = 0 Then
            PercentChange(ii) = 0
        Else
            PercentChange(ii) = YearlyChange(ii) / OpeningPrice_year(ii) 'cell format to be converted to %
        End If
        TotalVolume(ii) = Application.WorksheetFunction.Sum(Range(Cells(stockstartline(ii), 7), Cells(stockendline(ii), 7)))
    'copy the output to the o/p table
        Cells(ii + 2, 9) = TickerSymbol(ii)
        Cells(ii + 2, 10) = YearlyChange(ii)
        Cells(ii + 2, 11) = PercentChange(ii)
        Cells(ii + 2, 12) = TotalVolume(ii)
        If Cells(ii + 2, 11).Value > 0 Then
            Cells(ii + 2, 11).Interior.ColorIndex = 4
        Else: Cells(ii + 2, 11).Interior.ColorIndex = 3
        End If
    Next ii
    Range(Cells(2, 11), Cells(stock_count + 1, 11)).NumberFormat = "0.00%"
    greatestincrease = Application.WorksheetFunction.Max(PercentChange)
    Cells(2, 17) = greatestincrease
    Indx_GI = Application.WorksheetFunction.Match(greatestincrease, Range("k:k"), 0) 'provide row number
    Cells(2, 16) = TickerSymbol(Indx_GI - 2)
    Cells(2, 17).NumberFormat = "0.00%"
    greatestdecrease = Application.WorksheetFunction.Min(PercentChange)
    Indx_GI = WorksheetFunction.Match(greatestdecrease, Range("k:k"), 0) 'provide row number
    Cells(3, 16) = TickerSymbol(Indx_GI - 2)
    Cells(3, 17) = greatestdecrease
    Cells(3, 17).NumberFormat = "0.00%"
    greatestvol = Application.WorksheetFunction.Max(TotalVolume)
    Indx_GI = WorksheetFunction.Match(greatestvol, Range("l:l"), 0) 'provide row number
    Cells(4, 16) = TickerSymbol(Indx_GI - 2)
    Cells(4, 17) = greatestvol
    Cells(4, 17).NumberFormat = "general"
    Erase PercentChange 'to debug error when code executing third sheet
Next
End Sub

