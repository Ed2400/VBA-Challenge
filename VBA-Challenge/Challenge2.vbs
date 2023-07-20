Sub Challenge2()

'Definition of variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickers As Object
    Dim tickerCell As Range
    Dim dateCell As Range
    Dim openingCell As Range
    Dim closingCell As Range
    Dim yearlyChange As Double
    Dim totalVolume As Double
    Dim maxPercentageIncrease As Double
    Dim maxPercentageDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop in each ws
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Create a dictionary
        Set tickers = CreateObject("Scripting.Dictionary")
        
        ' Loop through each cell in columns A, B, C, F
    
        For Each tickerCell In ws.Range("A2:A" & lastRow)
        ' Offset by # of columns needed to get the corresponding data
            Set dateCell = tickerCell.Offset(0, 1)
            Set openingCell = tickerCell.Offset(0, 2)
            Set closingCell = tickerCell.Offset(0, 5)
            
            
            Dim year As Long
            year = Left(dateCell.Value, 4)
            
            ' If not is used as a way of discriminating the ones we already used.
            If Not tickers.Exists(tickerCell.Value) Then
             
                Set tickers(tickerCell.Value) = CreateObject("Scripting.Dictionary")
                tickers(tickerCell.Value)("Year") = year
                tickers(tickerCell.Value)("OpeningPrice") = openingCell.Value
            End If
            
            tickers(tickerCell.Value)("ClosingPrice") = closingCell.Value
            tickers(tickerCell.Value)("TotalVolume") = tickers(tickerCell.Value)("TotalVolume") + tickerCell.Offset(0, 6).Value
        Next tickerCell
        
        ' Write extracted data to the new columns
        ws.Range("M:P").ClearContents ' Clear previous data in columns M to P
        ws.Range("M1").Value = "Ticker"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Percentage Change"
        ws.Range("P1").Value = "Total Volume"
        
        Dim rowIndex As Long
        rowIndex = 2
        For Each Key In tickers.keys
            ' Calculate the yearly change
            yearlyChange = tickers(Key)("ClosingPrice") - tickers(Key)("OpeningPrice")
            
            ' Calculate the percentage change
            Dim percentageChange As Double
            If tickers(Key)("OpeningPrice") <> 0 Then
                percentageChange = yearlyChange / tickers(Key)("OpeningPrice") * 100
            Else
                percentageChange = 0
            End If

            ws.Cells(rowIndex, "M").Value = Key
            ws.Cells(rowIndex, "N").Value = yearlyChange
            ws.Cells(rowIndex, "O").Value = percentageChange
            ws.Cells(rowIndex, "P").Value = tickers(Key)("TotalVolume")
            
           
            If percentageChange > maxPercentageIncrease Then
                maxPercentageIncrease = percentageChange
                maxIncreaseTicker = Key
            End If
            
            If percentageChange < maxPercentageDecrease Then
                maxPercentageDecrease = percentageChange
                maxDecreaseTicker = Key
            End If
            
            If tickers(Key)("TotalVolume") > maxTotalVolume Then
                maxTotalVolume = tickers(Key)("TotalVolume")
                maxVolumeTicker = Key
            End If
            
            rowIndex = rowIndex + 1
        Next Key
        
        ' Printing the table
    
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S2").Value = maxIncreaseTicker
        ws.Range("S3").Value = maxDecreaseTicker
        ws.Range("S4").Value = maxVolumeTicker
        ws.Range("T2").Value = maxPercentageIncrease & "%"
        ws.Range("T3").Value = maxPercentageDecrease & "%"
        ws.Range("T4").Value = maxTotalVolume