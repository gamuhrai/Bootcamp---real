Attribute VB_Name = "Module1"
Sub StockTickerScanner()
    Application.ScreenUpdating = False

    Dim rowNum As Long
    Dim tickerRow As Long
    Dim total As Double
    Dim volStart As Long
    Dim gpincrease As Variant
    Dim gpdecrease As Variant
    Dim gtvol As Variant

    'application.screenupdating = false/true

    'writes column headers and labels for each sheet
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Since each has different number of rows we need
        ' a way to return total rows in each sheet

        rowNum = Cells(Rows.Count, "A").End(xlUp).Row

        total = 0
        tickerRow = 1
        volStart = 2
        openPrice1 = Cells(2, 3)

        For i = 2 To rowNum
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
                tickerRow = tickerRow + 1

                'prints name of ticker in column i
                ws.Cells(tickerRow, "I") = ws.Cells(i, "A")

                'calculate yearly change
                ychange = ws.Cells(i, 6) - openPrice1

                'prints year change in column j
                ws.Cells(tickerRow, "J") = ychange
                openPrice1 = ws.Cells(i + 1, "C")

                'calculate and print percent change
                If openPrice1 <> 0 Then
                    pchange = ychange / openPrice1
                    ws.Cells(tickerRow, "K") = pchange
                    ws.Cells(tickerRow, "K").NumberFormat = "0.00%"
                Else
                    ws.Cells(tickerRow, "K") = "N/A"
                End If

                'calculate total volume
                For j = volStart To i
                    total = total + ws.Cells(j, 7).Value
                Next j
                ws.Cells(tickerRow, "L") = total

                ' reset total for the new ticker
                total = 0
                volStart = i + 1

                'conditional formatting
                Select Case ychange
                    Case Is > 0
                        ws.Range("J" & tickerRow).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & tickerRow).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & tickerRow).Interior.ColorIndex = 0
                End Select
            End If
        Next i

        ' prints max, min and greatest volume to column Q
        ws.Cells("2", "Q") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowNum)) * 100
        ws.Cells("3", "Q") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowNum)) * 100
        ws.Cells("4", "Q") = WorksheetFunction.Max(ws.Range("L2:L" & rowNum))

        'need variables to hold the row info using match function
        'WorksheetFunction.Match(lookup_value, lookup_range, [match_type])

        gpincrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowNum)), ws.Range("K2:K" & rowNum), 0)
        gpdecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowNum)), ws.Range("K2:K" & rowNum), 0)
        gtvol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowNum)), ws.Range("L2:L" & rowNum), 0)

        'print ticker using previous variable
        ws.Range("P2") = ws.Cells(gpincrease, 9)
        ws.Range("P3") = ws.Cells(gpdecrease, 9)
        ws.Range("P4") = ws.Cells(gtvol, 9)
    Next ws

    Application.ScreenUpdating = True
End Sub
