Attribute VB_Name = "Module1"
Sub StocksEva()
    ' Automate info all spreadsheets
    For Each ws In Worksheets

    ' Name Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
     ' Identify Variables and set baseline values
        Dim TickerName1 As String
        Dim EndRow As Long
        Dim TotalTickerVulm As Double
        TotalTickerVulm = 0
        Dim SummaryTablesRow As Long
        SummaryTableRow = 2
        Dim YearlyOpen1 As Double
        Dim YearlyClose1 As Double
        Dim YearlyChange As Double
        Dim PreviousValue As Long
        PreviousValue = 2
        Dim PercentChange As Double
        
       ' Identify Endrow
           EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                                    For i = 2 To EndRow
                                    
        ' Add Up  Ticker Volume
            TotalTickerVulm = TotalTickerVulm + ws.Cells(i, 7).Value
        ' Check If We Are Still Within The Same Ticker Name If It Is Not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          ' TickerName1 is
            TickerName1 = ws.Cells(i, 1).Value
        ' List Ticker Name In Summary - Table
                ws.Range("I" & SummaryTableRow).Value = TickerName1
        ' List Ticker Total Amount in Summary - Table
            ws.Range("L" & SummaryTableRow).Value = TotalTickerVulm
        ' Return ticker total to baseline
                TotalTickerVulm = 0
        'List Yearly Open, Yearly Close and Yearly Change
                YearlyOpen1 = ws.Range("C" & PreviousValue)
                YearlyClose1 = ws.Range("F" & i)
                YearlyChange = YearlyClose1 - YearlyOpen1
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
        ' Identify Percent Change
                If YearlyOpen1 = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen1 = ws.Range("C" & PreviousValue)
                    PercentChange = YearlyChange / YearlyOpen1
                End If
         ' Convert Total Double for % Symbol And Decimal Places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange
         ' Highlight Cells - Positive (Green) / Negative (Red)
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
          ' Add One To The Summary Table Row
                SummaryTableRow = SummaryTableRow + 1
                PreviousValue = i + 1
                End If
                
            Next i

    Next ws
    
End Sub



