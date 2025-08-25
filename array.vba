Dim ws As worksheet
Dim practice_fun() As Variant

start = 6
lastRow = Range("A" & Rows.Count).End(xlUp).Row

lastCol = Cells(5, 1).End(xlToRight).Column
colletter = Split(Cells(5, lastCol).Address, "$")(1)


'array fill up
'practice_fun = Range("A7:B" & lastRow).Value
practice_fun = Range("A6:" & colletter & lastRow).Value


Set wb = Workbooks.Add

    For i = 1 To UBound(practice_fun)
        ActiveSheet.Name = practice_fun(i, 1)
        ActiveSheet.Range("A1").Value = practice_fun(i, 1)
        ActiveSheet.Range("A2").Value = practice_fun(i, 2)
        Worksheets.Add
    Next i

End Sub
