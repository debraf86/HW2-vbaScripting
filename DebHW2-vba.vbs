Sub stock_values()

  ' Declare variables
  Dim column, row, numRow, totalStockVol, yearlyChange As Double
  Dim openingAmt, closingAmt, percentChg, closingRow As Double
  Dim percentChgStr, greatestIncTicker, greatestDecTicker, greatestVolTicker As String
  Dim greatestPerInc, greatestPerDec, greatestTotVol As Double
  Dim WS_Count As Integer
  
  ' Get number of worksheets
  WS_Count = ActiveWorkbook.Worksheets.Count
  
  For Each ws In Worksheets
  
    Dim WorksheetName As String
  
    ' get the worksheet name
    WorksheetName = ws.Name
  
    ' set initial values
    row = 2
    totalStockVol = 0
    yearlyChange = 0
    closingRow = 2
    openingRow = 2
    percentChg = 0
    greatestPerInc = 0
    greatestIncTicker = 0
    greatestPerDec = 0
    greatedDecTicker = 0
    greatestTotVol = 0
    greatestVolTicker = 0
    column = 1

    ' Get the first opening amount for the year. 
    openingAmt = ws.Cells(2, 3).Value

    ' Get the last row
    numRows = ws.Cells(Rows.Count, column).End(xlUp).row
    ' Set the column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Volume"
  
    ' Loop through rows in the column
    For i = 2 To numRows

        ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then

            ' Set the stock volume for the year
            ws.Range("I" & row).Value = Cells(i, column).Value
            ws.Range("L" & row).Value = totalStockVol

            ' Determine the yearly change from opening to closing amount
            closingAmt = ws.Range("F" & closingRow).Value
            openingAmt = ws.Range("C" & openingRow).Value
            ChangeAMt = closingAmt - openingAmt
    
            ws.Range("J" & row).Value = ChangeAMt
            If ChangeAMt >= 0 Then
                ws.Range("J" & row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & row).Interior.ColorIndex = 3
            End If

            ' Determine the percentage of change from opening to closing. Check for dividing by 0.
            If openingAmt <> 0 Then
                percentChg = (ChangeAMt / openingAmt) * 100
                ' i should have used the Percent funtion instead
                ' as in something like Range("K" & row).Value = Format(ChangeAmt / openingAmt, "Percent")
                ' but haven't tested it.
                ws.Range("K" & row).Value = Str(percentChg) + "%"
            Else
                ws.Range("K" & row).Value = "0%"
            End If
 
            ' Find the latest % increase, %decrease, and total volumes of stocks
            If percentChg > greatestPerInc Then
                greatestPerInc = percentChg
                greatestIncTicker = ws.Range("I" & row).Value
            End If
        
            If percentChg < greatestPerDec Then
                greatestPerDec = percentChg
                greatestDecTicker = ws.Range("I" & row).Value
            End If
        
            If totalStockVol > greatestTotVol Then
                greatestTotVol = totalStockVol
                greatestVolTicker = ws.Range("I" & row).Value
            End If
        
            ' Next row
            row = row + 1
            openingRow = closingRow + 1
        
            ' Reset back to 0 for the next ticker's total stock volume
            totalStockVol = 0

        End If

        ' Add the next stock volumes to the total
        totalStockVol = totalStockVol + ws.Cells(i + 1, "G").Value
        closingRow = closingRow + 1
    
    Next i
  
    ' Set the values for the greatest % increase, % decrease, and total volume of stock values. 
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P2").Value = greatestIncTicker
    ws.Range("P3").Value = greatestDecTicker
    ws.Range("P4").Value = greatestVolTicker
    ws.Range("Q2").Value = greatestPerInc
    ws.Range("Q3").Value = greatestPerDec
    ws.Range("Q4").Value = greatestTotVol
  
  ' next worksheet
  Next ws

End Sub