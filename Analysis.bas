Attribute VB_Name = "Module1"
Sub Analysis()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim Volume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim LastRow As Long
    Dim SummaryRow As Integer
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestTickerIncrease As String
    Dim GreatestTickerDecrease As String
    Dim GreatestTickerVolume As String

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets

        ws.Activate
        
        ' Initialize variables
        Volume = 0
        SummaryRow = 2
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set up summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"

        ' Loop through rows
        OpenPrice = ws.Cells(2, 3).Value
        For i = 2 To LastRow
            ' Check if ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = QuarterlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                Volume = Volume + ws.Cells(i, 7).Value

                ' Output to summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = Volume
                ws.Cells(SummaryRow, 11).Value = QuarterlyChange
                ws.Cells(SummaryRow, 12).Value = PercentChange

                ' Format percent change
                ws.Cells(SummaryRow, 12).NumberFormat = "0.00%"

                ' Reset for the next ticker
                SummaryRow = SummaryRow + 1
                Volume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Accumulate volume
                Volume = Volume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Conditional formatting
        Dim FormatRange As Range
        Set FormatRange = ws.Range("K2:K" & SummaryRow - 1)
        FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        FormatRange.FormatConditions(FormatRange.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
        
        Set FormatRange = ws.Range("K2:K" & SummaryRow - 1)
        FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        FormatRange.FormatConditions(FormatRange.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)

        ' Calculate greatest values
        GreatestIncrease = WorksheetFunction.Max(ws.Range("L2:L" & SummaryRow - 1))
        GreatestDecrease = WorksheetFunction.Min(ws.Range("L2:L" & SummaryRow - 1))
        GreatestVolume = WorksheetFunction.Max(ws.Range("J2:J" & SummaryRow - 1))
        
        ' Find associated tickers
        GreatestTickerIncrease = ws.Cells(WorksheetFunction.Match(GreatestIncrease, ws.Range("L2:L" & SummaryRow - 1), 0) + 1, 9).Value
        GreatestTickerDecrease = ws.Cells(WorksheetFunction.Match(GreatestDecrease, ws.Range("L2:L" & SummaryRow - 1), 0) + 1, 9).Value
        GreatestTickerVolume = ws.Cells(WorksheetFunction.Match(GreatestVolume, ws.Range("J2:J" & SummaryRow - 1), 0) + 1, 9).Value
        
        ' Output greatest values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Volume"
        ws.Cells(2, 15).Value = GreatestTickerIncrease
        ws.Cells(3, 15).Value = GreatestTickerDecrease
        ws.Cells(4, 15).Value = GreatestTickerVolume
        ws.Cells(2, 16).Value = GreatestIncrease
        ws.Cells(3, 16).Value = GreatestDecrease
        ws.Cells(4, 16).Value = GreatestVolume

    Next ws

    MsgBox "Analysis complete!"


End Sub
