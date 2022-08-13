Attribute VB_Name = "vba_challenge"
Sub vba_challenge():

    Dim i, j As Integer
    Dim flag As Boolean
    Dim startVal, thisTickerCount As Double
    Dim nRows As Long
    Dim nTickers As Long
    Dim currentTicker As String
    Dim tickersCount As Long
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim maxInc, minInc, maxVol As Double
    Dim maxTic, minTic, volTic As String
    
    
    
    For Each ws In Worksheets
    startVal = ws.Cells(2, 3).Value
    currentTicker = ws.Cells(2, 1).Value
    nRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    thisTickerCount = 0
    
    maxInc = -9999
    minInc = 9999
    maxVol = 0
    
    tickersCount = 1
    
    For i = 2 To (nRows)
        flag = False
        thisTickerCount = thisTickerCount + ws.Cells(i, 7).Value

        If ws.Cells(i + 1, 1).Value <> currentTicker Then
            'Update this ticker's data
            nTickers = ws.Cells(Rows.Count, 9).End(xlUp).Row
            tickersCount = tickersCount + 1

            ws.Cells(tickersCount, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(tickersCount, 10).Value = ws.Cells(i, 6).Value - startVal
            ws.Cells(tickersCount, 11).Value = FormatPercent(Cells(tickersCount, 10).Value / startVal)
            ws.Cells(tickersCount, 12).Value = thisTickerCount
            'Update Formatting
            If Cells(tickersCount, 10).Value > 0 Then
                ws.Cells(tickersCount, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(tickersCount, 10).Interior.ColorIndex = 3
            End If
            'Update max/min vals
            If ws.Cells(tickersCount, 11).Value > maxInc Then
                maxInc = ws.Cells(tickersCount, 11).Value
                maxTic = ws.Cells(tickersCount, 9).Value
            End If
            If ws.Cells(tickersCount, 11).Value < minInc Then
                minInc = ws.Cells(tickersCount, 11).Value
                minTic = ws.Cells(tickersCount, 9).Value
            End If
            If ws.Cells(tickersCount, 12).Value > maxVol Then
                maxVol = ws.Cells(tickersCount, 12).Value
                volTic = ws.Cells(tickersCount, 9).Value
            End If
            'Initialise currentTicker and startVal for new ticker
            currentTicker = ws.Cells(i + 1, 1).Value
            thisTickerCount = ws.Cells(i + 1, 7).Value
            startVal = ws.Cells(i + 1, 3).Value
        End If
    
    Next i
    
    'BONUS: Output greatest percentages and volume
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P2").Value = maxTic
    ws.Range("P3").Value = minTic
    ws.Range("P4").Value = volTic
    ws.Range("Q2").Value = FormatPercent(maxInc)
    ws.Range("Q3").Value = FormatPercent(minInc)
    ws.Range("Q4").Value = maxVol
    
    Next ws
    
    
    
    
    
    
    
End Sub
