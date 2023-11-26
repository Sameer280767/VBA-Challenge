Attribute VB_Name = "Module1"
Sub stock_practice1()

'To loop all worksheets in the Workbook
For Each ws In Worksheets

'Declaring variables and defining valus for fixed referenes
Dim Stock_Name As String
Dim Stock_Total As Double
Dim Stock_Total_Row As Integer
Dim annual_opening As Double
Dim annual_closing As Double
Dim percent_increase As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Stock_Total = 0

Stock_Table_Row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Going through duplicate records and grouping stocks and summing volume. Also isolating opening and closing price for each stock

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Stock_Name = ws.Cells(i, 1).Value
Stock_Total = Stock_Total + Cells(i, 7).Value

annual_closing = ws.Cells(i, 6).Value

percent_increase = (annual_closing - annual_opening) / annual_opening

ws.Range("J" & Stock_Table_Row) = annual_closing - annual_opening
ws.Range("I" & Stock_Table_Row).Value = Stock_Name
ws.Range("L" & Stock_Table_Row).Value = Stock_Total
ws.Range("K" & Stock_Table_Row).Value = FormatPercent(percent_increase)


Stock_Table_Row = Stock_Table_Row + 1

Stock_Total = 0
Else

Stock_Total = Stock_Total + ws.Cells(i, 7).Value

If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

annual_opening = ws.Cells(i, 3).Value


End If

End If

Next i

'Conditionally formatting Yealy Change

LastRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To LastRow1

If ws.Cells(i, 10) < 0 Then

ws.Cells(i, 10).Interior.ColorIndex = 3

Else

ws.Cells(i, 10).Interior.ColorIndex = 4

End If

Next i

'Defining Greatest % increase, decrease and total volume

Greatest_Per_Inc = WorksheetFunction.Max(ws.Range("K:K"))
Greatest_Per_Dec = WorksheetFunction.Min(ws.Range("K:K"))
Greatest_Tot_Volume = WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("Q2").Value = FormatPercent(Greatest_Per_Inc)
ws.Range("Q3").Value = FormatPercent(Greatest_Per_Dec)
ws.Range("Q4").Value = Greatest_Tot_Volume

'Identifying stock code for above parameter

For i = 2 To LastRow1

If Greatest_Tot_Volume = ws.Cells(i, 12).Value Then

ws.Range("P4").Value = ws.Cells(i, 9).Value


ElseIf Greatest_Per_Inc = ws.Cells(i, 11) Then

ws.Range("P2").Value = ws.Cells(i, 9).Value


ElseIf Greatest_Per_Dec = ws.Cells(i, 11).Value Then

ws.Range("P3").Value = ws.Cells(i, 9).Value



End If

Next i





ws.Cells.EntireColumn.AutoFit


Next ws


End Sub
