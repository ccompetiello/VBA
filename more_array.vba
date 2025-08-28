Sub last()

Dim lastrow As Long
Dim lastcol As Long
Dim colletter As String
Dim arry As Variant


lastrow = Range("A" & Rows.Count).End(xlUp).Row

lastcol = Cells(5, Columns.Count).End(xlToLeft).Column

colletter = Split(Cells(5, lastcol).Address, "$")(1)

arry = Range("A5:" & colletter & lastrow).Value

'perform an operation (add 30 to the values in the 7th column)

 For i = LBound(arry, 1) + 1 To UBound(arry, 1)
        arry(i, 7) = arry(i, 7) + 30
    Next i

'write back to different spot

Range("J5").Resize(UBound(arry, 1), UBound(arry, 2)).Value = arry

End Sub
