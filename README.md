# VBA-Challenge
For Week 2 Class Homework
Sub Ticker()

For Each ws In Worksheets

Dim TickerName As String

Dim TotalVolume As Double

TotalVolume = 0

Dim TickerSummaryRow As Integer

TickerSummaryRow = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 12).Value = "TotalVolume"

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

TickerName = ws.Cells(i, 1).Value

TotalVolume = TotalVolume + ws.Cells(i, 7)

ws.Range("I" & TickerSummaryRow).Value = TickerName

ws.Range("L" & TickerSummaryRow).Value = TotalVolume

TickerSummaryRow = TickerSummaryRow + 1

TotalVolume = 0

Else

TotalVolume = TotalVolume + ws.Cells(i, 7)

End If

Next i

Next ws

End Sub

Sub YearlyChange()

For Each ws In Worksheets

Dim TickerName As String

Dim YearlyChange As Double

YearlyChange = 0

OpenNumber = 2

Dim TickerSummaryRow As Integer

TickerSummaryRow = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 10).Value = "YearlyChange"

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

YearlyChange = ws.Cells(i, 6).Value - Cells(OpenNumber, 3)

ws.Range("J" & TickerSummaryRow).Value = YearlyChange

TickerSummaryRow = TickerSummaryRow + 1

OpenNumber = i + 1

YearlyChange = 0

End If

Next i

Next ws

End Sub

Sub Color()

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i, 10).Value > 0 Then

ws.Cells(i, 10).Interior.ColorIndex = 4

Else

ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

Next ws

End Sub

Sub PercentageChange()

For Each ws In Worksheets

Dim TickerName As String

Dim PercentageChange As Double

OpenNumber = 2

Dim TickerSummaryRow As Integer

TickerSummaryRow = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 11).Value = "PercentageChange"

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

If ws.Cells(i, 6).Value - Cells(OpenNumber, 3).Value = 0 Or Cells(OpenNumber, 3).Value = 0 Then

PercentageChange = 0

Else

PercentageChange = (ws.Cells(i, 6).Value - Cells(OpenNumber, 3).Value) / Cells(OpenNumber, 3).Value

End If

ws.Range("K" & TickerSummaryRow).Value = (PercentageChange) * 100 & "%"

TickerSummaryRow = TickerSummaryRow + 1

OpenNumber = i + 1

PercentageChange = 0

End If

Next i

Next ws

End Sub
