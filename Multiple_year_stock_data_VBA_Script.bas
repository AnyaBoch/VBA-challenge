Attribute VB_Name = "Module1"
Sub TickerReport()
'declaring variables
Dim ws As Worksheet
Dim rowNum As Long
Dim outputRowNum As Long
Dim tickName As Variant
Dim lastRow As Long
Dim totalVol As Double

'more variables for our output table1
Dim openPrice As Double
Dim closePrice As Double
Dim changePrice As Double
Dim percentChPrice As Double

'making headers for all worksheets in the workbook
For Each ws In ThisWorkbook.Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Value"

'initialization
outputRowNum = 2
'tickName = Range("A2").Value
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
totalVol = 0
openPrice = ws.Cells(2, 3).Value

'loops through rows
 For rowNum = 2 To lastRow
 If ws.Cells(rowNum + 1, 1).Value <> ws.Cells(rowNum, 1).Value Then
 
 'add ticker name to the output table
 tickName = ws.Cells(rowNum, 1).Value
 ws.Range("I" & outputRowNum).Value = tickName
 
 'adding ticker volume total and printing the total in output table
totalVol = totalVol + ws.Cells(rowNum, 7).Value
ws.Range("L" & outputRowNum).Value = totalVol
ws.Range("L" & outputRowNum).NumberFormat = "$#,##0" 'money format

'calculating and printing quarterly change and percent change into table1
closePrice = ws.Cells(rowNum, 6).Value
changePrice = closePrice - openPrice
ws.Range("J" & outputRowNum).Value = changePrice
ws.Range("J" & outputRowNum).NumberFormat = "0.00"

'applying colors to the table1 based on change price
If ws.Range("J" & outputRowNum).Value > 0 Then
   ws.Cells(outputRowNum, 10).Interior.Color = RGB(74, 156, 67) ' Green for positive values
ElseIf ws.Range("J" & outputRowNum).Value < 0 Then
          ws.Cells(outputRowNum, 10).Interior.Color = RGB(247, 84, 103) ' Red for negative values
Else
    ws.Cells(outputRowNum, 10).Interior.Color = RGB(255, 255, 255) ' Optional: white for zero values
End If

'calculating percent change
percentChPrice = Round((changePrice / openPrice), 4)
ws.Range("K" & outputRowNum) = percentChPrice
ws.Range("K" & outputRowNum).NumberFormat = "0.00%"

 
 'reset total volume for next ticker
totalVol = 0

'printing next ticker for output
outputRowNum = outputRowNum + 1
openPrice = ws.Cells(rowNum + 1, 3).Value
closePrice = 0
 
 Else
 totalVol = totalVol + ws.Cells(rowNum, 7).Value

 End If

 Next rowNum 'end loop
 
''/////////////"Greatest % increase", "Greatest % decrease", and "Greatest total volume"


'more variables to find greatest increase, decrease and total volume///////////

Dim maxValue As Double
Dim minValue As Double
Dim maxtotalVol As Double
Dim maxRow As Long
Dim minRow As Long
''''''''''''''''''''''''''some new Headers for output table2
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

''''initializing max and min values
lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
maxValue = ws.Range("K2").Value
minValue = ws.Range("K2").Value
maxtotalVol = ws.Range("L2").Value

maxRow = 2
minRow = 2

''''new loop to find MAXs and MIN and MAX_VOL
For outputRowNum = 2 To lastRow

'MAX(COLUMN "K"(11))
 If ws.Cells(outputRowNum, 11).Value > maxValue Then
            maxRow = outputRowNum
            maxValue = ws.Range("K" & maxRow).Value
            ws.Range("Q2").Value = maxValue
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("P2").Value = ws.Range("I" & maxRow).Value 'printing names of ticker for MAX values
        End If
'MIN (COLUMN "K"(11))
If ws.Cells(outputRowNum, 11).Value < minValue Then
           minRow = outputRowNum
           minValue = ws.Cells(minRow, "K").Value
           ws.Range("Q3").Value = minValue
           ws.Range("Q3").NumberFormat = "0.00%"
           ws.Range("P3").Value = Range("I" & minRow).Value 'printing names of ticker for MIN values:
        End If
'MAX_VOL (COLUMN"L"(12))
If ws.Cells(outputRowNum, 12) > maxtotalVol Then
            maxtotalVol = ws.Cells(outputRowNum, 12)
            ws.Range("Q4").Value = maxtotalVol
            ws.Range("P4").Value = ws.Range("I" & outputRowNum).Value
            ws.Range("P4").NumberFormat = "$#,##0" 'money format
    End If
    
    Next outputRowNum
           
           Next ws

End Sub



