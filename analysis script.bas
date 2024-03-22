Attribute VB_Name = "Module1"
Sub analysis()
'Declearing variables
Dim ws As Worksheet
Dim lastRow As Long
Dim ticker As String
Dim summaryRow As Long
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalVolume As Double
Dim greatestPercentageIncrease As Double
Dim greatestPercentageDecrease As Double
Dim greatestTotalVolume As Double
Dim greatestPercentageIncreaseTicker As String
Dim greatestPercentageDecreaseTicker As String
Dim greatestTotalVolumeTicker As String




'Looping through the worksheets
For Each ws In ThisWorkbook.Worksheets
'Getting the last row of the worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'Creating Headers for the new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
'Getting the Yearly change, percentagechange, total stock volume and tickers associated with them
        ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
        summaryRow = 2
        openPrice = ws.[C2]
        totalVolume = 0
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(summaryRow, 12).Value = totalVolume
                totalVolume = 0
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & summaryRow).Value = ticker
                closePrice = ws.Cells(i, 6).Value

                yearlyChange = closePrice - openPrice
                ws.Cells(summaryRow, 10).Value = yearlyChange

                If openPrice = 0 Then
                ws.Cells(summaryRow, 11).Value = "NA"

                Else
                percentageChange = (closePrice - openPrice) / openPrice
                ws.Cells(summaryRow, 11).Value = percentageChange
            
                End If
                
                openPrice = ws.Cells(i + 1, 3).Value
                
                summaryRow = summaryRow + 1
                
                Else
                
                End If
             Next i
'Formatting the Yearly Change and Percentage change columns
             Dim lastrowcount2 As Long
             lastrowcount2 = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
             ws.Range("K2:K" & lastrowcount2).NumberFormat = "0.00%"
             
             For j = 2 To lastrowcount2
                If ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
                Else
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
                End If
             Next j
'Getting the Greatest % Increase Value and Ticker
greatestPercentageIncrease = 0.001
For K = 2 To lastrowcount2
    If (ws.Cells(K, 11).Value <> "NA") Then
        If (ws.Cells(K, 11).Value > greatestPercentageIncrease) Then
            greatestPercentageIncrease = ws.Cells(K, 11).Value
            greatestPercentageIncreaseTicker = ws.Cells(K, 9).Value
        End If
        ElseIf (ws.Cells(K, 11).Value = "NA") Then
        End If
Next K

'Assigning the value and Ticker to the columns
ws.[P2] = greatestPercentageIncreaseTicker
ws.[Q2] = greatestPercentageIncrease

'Getting the Greatest % Decrease Value and Ticker
greatestPercentageDecrease = 0
For m = 2 To lastrowcount2
    If (ws.Cells(m, 11).Value < greatestPercentageDecrease) Then
        greatest_percentage_decrease = ws.Cells(m, 11).Value
        greatestPercentageDecreaseTicker = ws.Cells(m, 9).Value
    End If
Next m

'Assigning the Value and Ticker to columns
ws.[P3] = greatestPercentageDecreaseTicker
ws.[Q3] = greatest_percentage_decrease
    
'Getting the Greatest Total Volume Value and Ticker

greatestTotalVolume = 1
For v = 2 To lastrowcount2
    If (ws.Cells(v, 12).Value > greatestTotalVolume) Then
        greatestTotalVolume = ws.Cells(v, 12).Value
        greatestTotalVolumeTicker = ws.Cells(v, 9).Value
    End If
Next v

'Assigning the Value and Ticker to the columns
ws.[P4] = greatestTotalVolumeTicker
ws.[Q4] = greatestTotalVolume
Next ws


End Sub
