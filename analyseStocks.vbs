' AnalyseStocks
' Goes through stocks and summarize:
' Yearly change(opening begining of the year and closing end of the year)
' Percent change for the yearly change
' Total Stock Volume

Sub AnalyseStocks()

    ' Declaration
    Dim row As Long
    Dim rowCount As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim greatestTotalStockVolumeTicker As String
    Dim greatestTotalStockVolumeValue As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIncreaseValue As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestPercentDecreaseValue As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim nextRow As Long
    Dim totalStockVolume As Double
    Dim ws As Worksheet

    ' Go through each sheet
    For Each ws In Sheets

        ' Initialization
        nextRow = 2
        totalStockVolume = 0
        
        greatestTotalStockVolumeTicker = ""
        greatestTotalStockVolumeValue = 0
        greatestPercentIncreaseTicker = ""
        greatestPercentIncreaseValue = 0
        greatestPercentDecreaseTicker = ""
        greatestPercentDecreaseValue = 0
        

        ' Create column headers...
        ' ticket, yearly change, percent change, total stock volume
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        
        
        ' Format percent change column
        ws.Range("K2").EntireColumn.NumberFormat = "0.00%"
        
        ' Format yearly price change column
        ws.Range("J2").EntireColumn.NumberFormat = "0.00"
        ws.Range("J2").EntireColumn.Font.Color = vbBlack
        
        ' Format total stock volume colum
        ws.Range("L2").EntireColumn.NumberFormat = "#,##0.00"

        ' Get the row count
        rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

        ' Loop through all tickers data
        For row = 2 To rowCount

        ' Check if we are beginning a ticker block...
        If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
            totalStockVolume = 0
            ' TODO save the opening price
            openingPrice = ws.Cells(row, 3).Value
        End If

        ' For each row add to total stock volume
        totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value ' 7 is column number 7

        ' Check if we are ending a ticker block...
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            ' Ticker value
            ws.Cells(nextRow, 9).Value = ws.Cells(row, 1) ' 9 is column number - Ticker
            ' Save the closing price
            closingPrice = ws.Cells(row, 6).Value
            ' Compute yearly changes and store in column 10 - yearly change
            yearlyChange = closingPrice - openingPrice
            ws.Cells(nextRow, 10).Value = yearlyChange
             ' Color the yearly change cell - Red for negative, Green for positive
            If yearlyChange >= 0 Then
              ws.Cells(nextRow, 10).Interior.Color = RGB(0, 255, 0)
              Else
              ws.Cells(nextRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            ' Percent change closing - opening / opening
            If openingPrice <> 0 Then
            percentChange = ((closingPrice - openingPrice) / openingPrice)
            Else
            percentChange = (closingPrice - openingPrice)
            End If
            
            ' ws.Cells(nextRow, 11).Value = ((closingPrice - openingPrice) / 100)
            ws.Cells(nextRow, 12).Value = totalStockVolume ' 12 is column number - Total Stock Volume
            
            ' Second summary table...
            ' If totalStockVolume is greater than greatestTotalStockVolume then
            ' update greatestTotalStockVolumeTicker and greatestTotalStockVolumeValue...
            If totalStockVolume > greatestTotalStockVolumeValue Then
                greatestTotalStockVolumeValue = totalStockVolume
                greatestTotalStockVolumeTicker = ws.Cells(nextRow, 9).Value
            End If
            

            If percentChange > greatestPercentIncreaseValue Then
                greatestPercentIncreaseValue = percentChange
                greatestPercentIncreaseTicker = ws.Cells(nextRow, 9).Value
            End If
            
            If percentChange < greatestPercentDecreaseValue Then
                greatestPercentDecreaseValue = percentChange
                greatestPercentDecreaseTicker = ws.Cells(nextRow, 9).Value
            End If
           
            nextRow = nextRow + 1
        End If

        Next row
        ' Create secondary table...
        ' Greatest % columns headers
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % increase"
        ws.Range("O3") = "Greatest % decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        ' Format greatestPercentIncreaseValue
        'ws.Range("Q2").EntireColumn.NumberFormat = "0.00%"
        
        ' Format greatestTotalStockVolumeValue
        ws.Range("Q4").EntireColumn.NumberFormat = "#,##0.00"
        
        ws.Cells(2, 16).Value = greatestPercentIncreaseTicker
        ws.Cells(2, 17).Value = greatestPercentIncreaseValue
         ws.Cells(3, 16).Value = greatestPercentDecreaseTicker
        ws.Cells(3, 17).Value = greatestPercentDecreaseValue
        ws.Cells(4, 16).Value = greatestTotalStockVolumeTicker
        ws.Cells(4, 17).Value = greatestTotalStockVolumeValue
        
        Next ws

End Sub