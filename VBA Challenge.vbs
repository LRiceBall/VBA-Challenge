Sub Stock_Market_Test()
For Each ws In Worksheets

' Set summary table
ws.Range("I1, O1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"



    ' Creat a script that will loop through all the stocks
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim TotalStock As Double
    TotalStock = 0
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim TickerName As String
    Dim SummaryTable As Integer
    SummaryTable = 2
    Dim YearlyChange As Double
    Dim PercentChange As Integer
    Dim Max as Double
 
     
    ' Loop through rows in the column
    ' Start at 2 because the data does
    For i = 2 To LastRow
        
        ' Searches for when the value of the next cell is different than that of the current cell
        ' to find multiple items in large data columns
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerName = ws.Cells(i, 1).Value
            OpenValue = ws.Cells(i, 3).Value
            CloseValue = ws.Cells(i, 6).Value
            YearlyChange = CloseValue - OpenValue
            PercentChange = OpenValue / CloseValue
            TotalStock = TotalStock + CloseValue
            ' create summary table with accumulated data in colums G and H
            ' print out all CC types in column G
            ws.Range("I" & SummaryTable).Value = TickerName
            ws.Range("J" & SummaryTable).Value = YearlyChange
            ws.Range("K" & SummaryTable).value = PercentChange
            ' print out total charges per cc type in column H
            ws.Range("L" & SummaryTable).Value = TotalStock
            SummaryTable = SummaryTable + 1


        End If

    Next i

    For i = 2 to LastRow
        If ws.cells(i, 10).value < 0 Then
            ws.cells(i,10).interior.colorindex = 3
        elseif ws.cells(i, 10).value >= 0 Then
            ws.cells(i,10).interior.colorindex = 4
        end If
    next i


    For i = 2 To LastRow
        if ws.cells(i,10).value > 0 then
            If ws.Cells(i, 11).value > Max Then 
                Max = ws.Cells(i, 11).value
                ws.range("P2").Value = Max
                TickerName = ws.cells(i,9).value
                ws.range("O2").value = Tickername
            End if
        end if
    Next i

    For i = 2 To LastRow
        if ws.cells(i,10).value < 0 then
            If ws.Cells(i, 11).value > Max Then 
                Max = ws.Cells(i, 11).value
                ws.range("P3").Value = Max
                TickerName = ws.cells(i,9).value
                ws.range("O3").value = Tickername
            End if
        end if
    Next i 

 
    For i = 2 To LastRow
        If ws.Cells(i, 12).value > Max Then 
            Max = ws.Cells(i, 12).value
            ws.range("P4").Value = Max
            TickerName = ws.cells(i,9).value
            ws.range("O4").value = Tickername
        End if
    Next i


Next ws

End Sub



