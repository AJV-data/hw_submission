Sub stocks2():
For Each ws In Worksheets
    Dim stock_name As String
    Dim vol_total As Double
    Dim sum_t_row As Integer
    Dim y_open As Double
    Dim y_close As Double
    Dim temp As Double
    Dim y_change As Double
    Dim y_per As Double
    
    sum_t_row = 2
    cc_total = 0
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    y_open = Cells(2, 3)
    For i = 2 To lastrow
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            y_close = Cells(i, 6)
            vol_total = vol_total + Cells(i, 7)
            stock_name = Cells(i, 1)
            Cells(sum_t_row, 9) = stock_name
            Cells(sum_t_row, 12) = vol_total
            temp = y_open
            y_change = temp - y_close
            If temp = 0 Then
                y_per = 0
            Else
                y_per = (y_change / temp)
            End If
            Cells(sum_t_row, 11) = y_per
            Cells(sum_t_row, 10) = y_change
            
            y_open = Cells(i + 1, 3)
            vol_total = 0
            stock_name = ""
            sum_t_row = sum_t_row + 1
            
        Else
            vol_total = vol_total + Cells(i, 7)
        End If
    Next i
Next ws
End Sub
    