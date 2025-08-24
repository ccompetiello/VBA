
Dim quant() As Variant
Dim lastRow As Long
Dim i As Long
Dim r As Long
    
    lastRow = Range("E" & Rows.Count).End(xlUp).Row
    quant = Range("E6:E" & lastRow).Value
    
'    For i = 1 To UBound(quant, 1)
'    Next i
    
    For r = 1 To UBound(quant)
    
    quant(r, 1) = quant(r, 1) + 10
    
    Next r
    
'write back to column

Range("G6:G" & lastRow).Value = quant

End Sub

