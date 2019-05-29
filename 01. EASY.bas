Attribute VB_Name = "Easy"
Sub Easy():

Dim tick_name As String
Dim tick_total As Double
Dim lastrow, sum_row As Integer

tick_total = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
sum_row = 2

    For i = 2 To lastrow

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            tick_name = Cells(i, 1).Value
            tick_total = tick_total + Cells(i, 7).Value
            
            Range("I" & sum_row).Value = tick_name
            Range("J" & sum_row).Value = tick_total
            sum_row = sum_row + 1

            tick_total = 0

        Else

            tick_total = tick_total + Cells(i, 7).Value
 
        End If

    Next i

    'formatting
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Tot. Vol."
    
    Range("I1:J1").Select
    Selection.Font.Bold = True
    Selection.Interior.Color = 65535
    Selection.HorizontalAlignment = xlCenter
    
    Columns("J:J").Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
    
    Range("A1").Select

End Sub
