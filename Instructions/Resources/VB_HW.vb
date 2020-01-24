Sub ticker():

Dim Total As Double
Dim Assignment As Worksheet
Dim J As Integer

For Each Assignment In Worksheets
J = 0
Total = 0


RowCount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount

    Assignment.Range("I1").Value = "Ticker"
    Assignment.Range("J1").Value = "Total"

    If Assignment.Cells(i + 1, 1).Value <> Assignment.Cells(i, 1).Value Then
        Total = Total + Assignment.Cells(i, 7).Value

           ' Print ticker symbol
        Assignment.Range("I" & 2 + J).Value = Cells(i, 1).Value
           ' Print total
        Assignment.Range("J" & 2 + J).Value = Total
           ' Reset Total
        Total = 0
           ' Move to next row
        J = J + 1

    Else
        Total = Total + Assignment.Cells(i, 7).Value
    End If

    Next i
    Next Assignment

End Sub
