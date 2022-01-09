Refactored Code
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

    Worksheets("All Stocks Analysis").Activate   
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
Worksheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        tickerIndex = 0

    For i = 0 To 11
       tickerVolumes(i) = 0
    Next i

For j = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

            If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
              tickerStartingPrices(tickerIndex) = Cells(j, 6).Value  

            End If
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(j, 6).Value   

            tickerIndex = tickerIndex + 1     

          End If
          
        Next j
         For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
Next i
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For j = dataRowStart To dataRowEnd    
        If Cells(j, 3) > 0 Then           
            Cells(j, 3).Interior.Color = vbGreen      
        Else 
            Cells(j, 3).Interior.Color = vbRed      
        End If
    Next j
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

Sub ClearWorksheet()

    Cells.Clear
    
End Sub
