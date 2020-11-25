Sub alphabet_testing()
 Dim i As Double
 Dim j As Integer
 Dim Start As Double
 Dim ticker_nm As String
 Dim total_vol As Double
 Dim Yearly_change As Double
 Dim Percent_change As Double
 Dim Open_Value As Double
 total_vol = 0
 i = 2
 j = 2
 Start = 2
 LastRow = Cells(Rows.Count, "A").End(xlUp).Row
 For i = 2 To LastRow
    total_vol = total_vol + Cells(i, 7).Value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_nm = Cells(i, 1).Value
        total_vol = total_vol + Cells(i, 7).Value
        Open_Value = Cells(Start, 3).Value
        Yearly_change = Cells(i, 6).Value - Open_Value
        If Yearly_change <> 0 And Open_Value <> 0 Then
            Percent_change = (Yearly_change / Cells(Start, 3).Value) * 100
            Else
                Percent_change = 0
        End If
        Range("I" & j) = ticker_nm
        Range("L" & j) = total_vol
        Range("J" & j) = Yearly_change
        Range("K" & j) = Percent_change
        total_vol = 0
        j = j + 1
        Start = i + 1
   End If
    Next i
End Sub

