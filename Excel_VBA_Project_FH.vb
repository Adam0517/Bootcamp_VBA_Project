Sub VBA_Project2_FH()

'Do each loop for each sheet

For Each ws In Worksheets
'Create header 
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total StockVolume"

'Need to chek the rows which cells are zero and then delete it.
'In order to avoid error when we do calculate "Percentage Change"
    Dim lastrow_zero As Double
    Dim zero_first As Double

    lastrow_zero = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Calculate how much rows we have in original sheet
    ws.Cells(lastrow_zero + 1, "C").Value = 0 'Add a zero value in last row+1 to avoid the error for below function
    zero_first = Split(ws.Range("C2:C" & (lastrow_zero + 1)).Find(what:=0).Address, "$")(2) 'Get the first zero in open price column

'To delete the row which cells are all of zero from last row to zero_first
    For j = (lastrow_zero + 1) To zero_first Step -1
            If ws.Cells(j, "C") = 0 Then
                ws.Cells(j, "C").EntireRow.Delete
            End If
    Next j

'Find how much categories of Tickers
    Dim myrange As Range
    Dim lastrow As Double
    Dim cate_lastrow As Double

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'This number is how much rows after deleting zero rows
    ws.Range("I2:I" & lastrow).Value = ws.Range("A2:A" & lastrow).Value

    Set myrange = ws.Range("I1:I" & lastrow)
    myrange.RemoveDuplicates Columns:=1, Header:=xlYes 'Use remove duplicate way to get categories

'Count how much cate (I column)
    cate_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Total of stock for each Ticker
    Dim total As Double
    Dim c_count As Double
    Dim t_count As Double
    Dim k_count As Double

'Initial values for t_count and k_count
    t_count = 0
    k_count = 2


    For i = 1 To (cate_lastrow - 1)
'Total stock : sum volumns by each category of Ticker
    total = WorksheetFunction.SumIf(ws.Range("A2:A" & lastrow), ws.Cells(i + 1, "I").Value, ws.Range("G2:G" & lastrow))
    ws.Cells(i + 1, "L").Value = total

'Yearly Change : First value of each category - last value of each category
    c_count = WorksheetFunction.CountIf(ws.Range("A2:A" & lastrow), ws.Cells(i + 1, "I").Value)
    t_count = t_count + c_count
    ws.Cells(i + 1, "J").Value = ws.Cells(1 + t_count, "F").Value - ws.Cells(k_count, "F").Value

'Percent Change 
    ws.Cells(i + 1, "K").Value = ws.Cells(i + 1, "J").Value / ws.Cells(k_count, "F").Value

    k_count = t_count + 2
    Next i

'Change the Percent Change into %
    ws.Range("K2:K" & cate_lastrow).NumberFormat = "0.00%"

'Conditional formating for yearly change : Green 4(>=0) and Red 3(<0)
    Dim gr As Range
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition

    Set gr = ws.Range("J2:J" & cate_lastrow)
    Set cond1 = gr.FormatConditions.Add(xlCellValue, xlLess, "0")
    Set cond2 = gr.FormatConditions.Add(xlCellValue, xlGreaterEqual, "0")

    cond1.Interior.ColorIndex = 3
    cond2.Interior.ColorIndex = 4

'Locate the stock with Greatest % Increase/Decrease and Total Volume
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    Dim com_index As String 'Record the max/min/total index
    Dim com_num As Double 'Record the max/min/total number

'Greatest % Increase
    ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & cate_lastrow))
    com_index = ws.Range("K2:K" & cate_lastrow).Find(what:=ws.Range("P2").Value * 100).Address
    com_num = Split(com_index, "$")(2)
    ws.Range("O2").Value = ws.Cells(com_num, "I").Value

'Greatest % Decrease
    ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K2:K" & cate_lastrow))
    com_index = ws.Range("K2:K" & cate_lastrow).Find(what:=ws.Range("P3").Value * 100).Address
    com_num = Split(com_index, "$")(2)
    ws.Range("O3").Value = ws.Cells(com_num, "I").Value
' Add % for P2 and P3
    ws.Range("P2:P3").NumberFormat = "0.00%"
'Greatest Total Volume
    ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & cate_lastrow))
    com_index = ws.Range("L2:L" & cate_lastrow).Find(what:=ws.Range("P4").Value).Address
    com_num = Split(com_index, "$")(2)
    ws.Range("O4").Value = ws.Cells(com_num, "I").Value

Next ws

End Sub




