Attribute VB_Name = "Module1"
Sub getTickerYearlySummary()


For Each ws In ThisWorkbook.Worksheets
    
    ' Initialize variables
    Dim startCellRow As Long
    Dim endCellRow As Long
    Dim summaryTableRow As Integer
    Dim lastRow As Long
    Dim lastRowSummary As Integer
    
    Dim volume As Double
    Dim tickerSymbol As String
    Dim yearOpenPrice As Double
    Dim yearClosePrice As Double
    Dim priceChange As Double
    Dim percentChange As Double
    Dim highestVolume As Double
    Dim highestVolumeSymbol As String
    Dim highestPercent As Double
    Dim highestPercentSymbol As String
    Dim lowestPercent As Double
    Dim lowestPercentSymbol As String
    
    
    ' Set starting values for variables
    startCellRow = 2
    volume = 0
    summaryTableRow = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create Column and Row Headings
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest %  Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    
    'Loop through all rows
    For i = 2 To lastRow

        
        'Check to see if ticker symbol has changed.  If it has then loop through all the rows for that symbol
        'and total the volume, get the opening and closing prices, and calculate the price and percentage change.
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            endCellRow = i
            yearOpenPrice = ws.Range("C" & startCellRow).Value
            yearClosePrice = ws.Range("F" & endCellRow).Value
            priceChange = yearClosePrice - yearOpenPrice
            
            If yearOpenPrice = 0 Or priceChange = 0 Then
                percentChange = 0
            Else
                percentChange = priceChange / yearOpenPrice
            End If

'            'MsgBox ("PriceChange=" & priceChange & ", PercentChange=" & percentChange)
            
            For j = startCellRow To endCellRow
                If j <= endCellRow Then
                    volume = volume + ws.Cells(j, 7)
                End If
            Next j
            
            tickerSymbol = ws.Cells(i, 1).Value
            ws.Range("I" & summaryTableRow).Value = tickerSymbol
            ws.Range("J" & summaryTableRow).Value = priceChange
            ws.Range("K" & summaryTableRow).Value = percentChange
            ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
            ws.Range("L" & summaryTableRow).Value = volume
            
            If (priceChange > 0) Then
                ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                
            ElseIf (priceChange < 0) Then
                ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3

            End If
                     
            'Increment startCellRow and summaryTableRow and reset volume to 0
            startCellRow = i + 1
            summaryTableRow = summaryTableRow + 1
            volume = 0

        End If
    
    Next i
    

    'Get the symbol with the highest volume and print symbol and volume
    
    lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    highestVolume = Application.WorksheetFunction.Max(ws.Range("l:l"))
    
    For x = 2 To lastRowSummary
        If ws.Range("L" & x).Value = highestVolume Then
            highestVolumeSymbol = ws.Range("I" & x).Value
            ws.Range("O4").Value = highestVolumeSymbol
            ws.Range("P4").Value = highestVolume
        End If
            
     Next x
     
     
    'Loop through the column with percentage change and find the highest and lowest percent increase
    'Print the symbol and values for highest and lowest percent increase

    
    highestPercent = Application.WorksheetFunction.Max(ws.Range("k:k"))
    lowestPercent = Application.WorksheetFunction.Min(ws.Range("k:k"))
    
    For y = 2 To lastRowSummary
        If ws.Range("K" & y).Value = highestPercent Then
            highestPercentSymbol = ws.Range("I" & y).Value
            ws.Range("O2").Value = highestPercentSymbol
            ws.Range("P2").Value = highestPercent
            ws.Range("P2").NumberFormat = "0.00%"
        
        ElseIf ws.Range("K" & y).Value = lowestPercent Then
            lowestPercentSymbol = ws.Range("I" & y).Value
            ws.Range("O3").Value = lowestPercentSymbol
            ws.Range("P3").Value = lowestPercent
            ws.Range("P3").NumberFormat = "0.00%"
        
        End If
            
     Next y
     
    
    Next
 
    'starting_ws.Activate
    MsgBox ("finished")

End Sub


Sub ClearCells()

'Dim lastRow As Long
'lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For Each ws In Worksheets
            ws.Range("I1:P1000").Clear
        Next ws
        
    MsgBox ("finished")
End Sub





