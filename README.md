Sub project2()
    For Each ws In Worksheets
    Dim thisRow As String
    Dim nextRow As String
    Dim tsv As LongLong
    tsv = 0
    Dim yearChange As Double
    Dim yearOpen As Double
    Dim newNum As Long
    Dim yearClose As Double
    Dim perChange As Double
    Dim gpIncrease As Double
    Dim gpDecrease As Double
    Dim gTSV As LongLong
    Dim gpiTicker As String
    Dim gpdTicker As String
    Dim gtsvTicker As String
    
    
    newNum = 2
    gpIncrease = 0
    gpDecrease = 0
    gTSV = 0
    
    Set curSheet = ThisWorkbook.Worksheets("2014")
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    yearOpen = ws.Cells(2, 3).Value
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        thisRow = ws.Cells(i, 1)
        nextRow = ws.Cells(i + 1, 1)
        If thisRow = nextRow Then
            tsv = tsv + ws.Cells(i, 7).Value

        Else
            
            yearClose = ws.Cells(i, 6).Value
            yearChange = yearClose - yearOpen
            If yearOpen = 0 Then
            perChange = 0
            Else
            perChange = (yearChange / yearOpen)
            End If
        
            
            
            tsv = tsv + curSheet.Cells(i, 7).Value
            ws.Cells(newNum, 12) = tsv
            ws.Cells(newNum, 9) = thisRow
            ws.Cells(newNum, 10) = yearChange
            ws.Cells(newNum, 11) = perChange
            If ws.Cells(newNum, 10) > 0 Then
            ws.Cells(newNum, 10).Interior.Color = RGB(0, 255, 0)
            Else
            ws.Cells(newNum, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            

            
            tsv = 0
            yearOpen = ws.Cells(i + 1, 3).Value
            newNum = newNum + 1
        End If
            ' greatest percent increase gpIncrease perChange
            If (perChange > gpIncrease) Then
            gpIncrease = perChange
            gpiTicker = thisRow
            End If
            
            ' greatest percent decrease gpDecrease perChange
            If (perChange < gpDecrease) Then
            gpDecrease = perChange
            gpdTicker = thisRow
            End If
            
            ' greatest total volume gTSV tsv
            If (tsv > gTSV) Then
            gTSV = tsv
            gtsvTicker = thisRow
            End If
    Next i
    ws.Cells(2, 17) = gpIncrease
    ws.Cells(3, 17) = gpDecrease
    ws.Cells(4, 17) = gTSV
    ws.Cells(2, 16) = gpiTicker
    ws.Cells(3, 16) = gpdTicker
    ws.Cells(4, 16) = gstvTicker
    ws.Range("K1:K" & lastRow).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
Next ws
End Sub


