' fill in for ?
' follow the hints and comments



Sub SolveStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim ticker As String
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String



    ' Loop through each sheet (Q1, Q2, Q3, Q4)
    ' Hint: Each ? In ThisWorkbook.?
    For ?

        ' Hint: Activate each ws
        ws.?

        ' Find the last row of data
        ' Hint: we had done a similar one during lecture
        ' Hint: .End(xlUp).Row
        lastRow = ?



        ' Initialize variables
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        startRow = 2


        ' Add headers for new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(1, 14).Value = "Metrics"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        Dim j As Integer
        j = 0



        ' Process each row
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value



            ' If the ticker changes or it is the last row
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then



                ' Calculate Quarterly Change and Percent Change
                ' Hint: Cells or Range
                If ws.?(startRow, 3).Value <> 0 Then
                    quarterlyChange = ws.?(i, 6).Value - ws.?(startRow, 3).Value
                    percentChange = quarterlyChange / ws.?(startRow, 3).Value
                Else
                    quarterlyChange = 0
                    percentChange = 0
                End If



                ' Write data to the output table
                ws.Cells(2 + j, 9).Value = ?  ' ticker
                ws.Cells(2 + j, 10).Value = ?  ' quarterly change
                ws.Cells(2 + j, 10).NumberFormat = "0.00"
                ws.Cells(2 + j, 11).Value = ?  ' percent change
                ws.Cells(2 + j, 11).NumberFormat = "0.00%"
                ws.Cells(2 + j, 12).Value = ?  ' total volume



                ' Conditional formatting for positive/negative changes
                ' Hint: Interior.ColorIndex
                ' Cells or Range
                If quarterlyChange > 0 Then
                    ws.?(2 + j, 10).? = 4 ' Green
                ElseIf quarterlyChange < 0 Then
                    ws.?(2 + j, 10).? = 3 ' Red
                Else
                    ws.?(2 + j, 10).? = 0 ' No color
                End If



                ' Track greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = ?
                    greatestIncreaseTicker = ?
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = ?
                    greatestDecreaseTicker = ?
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = ?
                    greatestVolumeTicker = ?
                End If

                ' Reset variables for next ticker
                totalVolume = 0
                startRow = i + 1
                j = j + 1
            End If
        Next i


        ' Output greatest values
        ws.Cells(2, 15).Value = ? ' greatest increase ticker
        ws.Cells(2, 16).Value = ? ' greatest increase
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = ? ' greatest decrease ticker
        ws.Cells(3, 16).Value = ? ' greatest decrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = ? ' greatest volume ticker
        ws.Cells(4, 16).Value = ? ' greatest volume

    Next ws

    MsgBox "Data processed successfully!"
End Sub
