Attribute VB_Name = "Module1"
Sub analyze()
    'Iterate through all sheets
    For Each ws In ActiveWorkbook.worksheets
        ws.Activate
        
        'Find last row with value
        Dim lastRow As Long
        lastRow = Cells(Rows.count, "A").End(xlUp).Row
        'ticker
        Dim ticker As String
        'Open Value
        Dim openValue As Double
        'Close Value
        Dim closeValue As Double
        'Total stock volume
        Dim totalVolume As LongLong
        'Summary Table row
        Dim tableRow As Integer
        tableRow = 2
        'Rows traversed with same ticker
        Dim numbRows As Integer
        numbRows = 1
        
            'Iterate through data
            For i = 2 To lastRow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    ticker = Cells(i, 1).Value
                    totalVolume = brandTotal + Cells(i, 7).Value
                    Range("I" & tableRow).Value = ticker
                    Range("L" & tableRow).Value = totalVolume
                    openValue = Cells(i - numbRows + 1, 3).Value
                    closeValue = Cells(i, 6).Value
                    Range("J" & tableRow).Value = openValue - closeValue
                    If openValue = 0 Then
                        Range("K" & tableRow).Value = 0 - closeValue
                    Else
                        Range("K" & tableRow).Value = (openValue - closeValue) / openValue
                    End If
                    tableRow = tableRow + 1
                    totalVolume = 0
                    numbRows = 0
                Else
                    totalVolume = totalVolume + Cells(i, 7).Value
                    numbRows = numbRows + 1
                End If
            Next i
            
            'Formatting
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            Range("O2").Value = "Greatest % increase"
            Range("O3").Value = "Greatest % decrease"
            Range("O4").Value = "Greatest total volume"
            Range("K:K").NumberFormat = "0.00%"
            'Conditional Formatting for Percent Change
            Dim rng As Range
            Dim cell As Range
            Set rng = Range(Range("K2"), Range("K2").End(xlDown))
            For Each cell In rng.Cells
                If cell.Value < 0 Then
                    cell.Interior.ColorIndex = 3
                ElseIf cell.Value > 0 Then
                    cell.Interior.ColorIndex = 4
                End If
            Next
            
            'Bonus "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Dim greatestIncrease As Double
            greatestIncrease = 0
            Dim greatestIncreaseTicker As String
            Dim greatestDecrease As Double
            greatestDecrease = 0
            Dim greatestDecreaseTicker As String
            Dim greatestTotalVolume As LongLong
            greatestTotalVolume = 0
            Dim greatestTotalVolumeTicker As String
            Dim lastPercentRow As Long
            lastPercentRow = Cells(Rows.count, "K").End(xlUp).Row
            For x = 2 To lastPercentRow
                If Range("K" & x).Value > greatestIncrease Then
                    greatestIncrease = Range("K" & x).Value
                    greatestIncreaseTicker = Range("I" & x).Value
                ElseIf Range("K" & x).Value < greatestDecrease Then
                    greatestDecrease = Range("K" & x).Value
                    greatestDecreaseTicker = Range("I" & x).Value
                ElseIf Range("L" & x).Value > greatestTotalVolume Then
                    greatestTotalVolume = Range("L" & x).Value
                    greatestTotalVolumeTicker = Range("I" & x).Value
                End If
            Next x
            Range("P2").Value = greatestIncreaseTicker
            Range("P3").Value = greatestDecreaseTicker
            Range("P4").Value = greatestTotalVolumeTicker
            Range("Q2").Value = greatestIncrease
            Range("Q3").Value = greatestDecrease
            Range("Q2:Q3").NumberFormat = "0.00%"
            Range("Q4").Value = greatestTotalVolume
            
            'Autofit worksheet
            Columns("A:P").AutoFit
            
    Next ws
End Sub
