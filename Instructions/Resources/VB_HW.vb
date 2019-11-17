Sub ticker():

Dim Total As Double

RowCount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Total = Total + Cells(i, 7).Value

           ' Print ticker symbol
        Range("I" & 2 + j).Value = Cells(i, 1).Value
           ' Print total
        Range("J" & 2 + j).Value = Total
           ' Reset Total
        Total = 0
           ' Move to next row
        j = j + 1

    Else
        Total = Total + Cells(i, 7).Value
    End If

    Next i


End Sub