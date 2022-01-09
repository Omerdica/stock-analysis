Original Code
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer    
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All stocks (" + yearValue + ")"
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"
Worksheets(yearValue).Activate
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'Loop through tickers
   For i = 0 To 11
       Ticker = tickers(i)
       totalVolume = 0

 For j = 2 To RowCount
      'Get total volume for current ticker
           If Cells(j, 1).Value = Ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If
           'get starting price for current ticker
           If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
               startingPrice = Cells(j, 6).Value
           End If
           If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
               endingPrice = Cells(j, 6).Value

           End If
       Next j
       'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = Ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


' Formating
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0$"
        Range("c4:C15").NumberFormat = "0,00%"
        Columns("B").AutoFit
        
        If Cells(4, 3) > 0 Then
            'Color the cells green
            Cells(4, 3).Interior.Color = vbGreen
        ElseIf Cells(4, 3) < 0 Then
        
            'Color the cell red
            Cells(4, 3).Interior.Color = vbRed
        Else
            'clear the cell color
                Cells(4, 3).Interior.Color = x1None
            
        End If
        
         dataRowStart = 4
        dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

        'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
End Sub

Sub ClearWorksheet()

    Cells.Clear
    
End Sub


    
   
