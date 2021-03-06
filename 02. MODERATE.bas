Attribute VB_Name = "Moderate"
Sub Moderate():

Dim tick_name As String
Dim tick_open, tick_close, tick_change, tick_perc, tick_total As Double
Dim lastrow, sum_row As Integer

tick_open = 0
tick_close = 0
tick_total = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
sum_row = 2

    For i = 2 To lastrow

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            tick_name = Cells(i, 1).Value
            tick_close = Cells(i, 6).Value
            tick_total = tick_total + Cells(i, 7).Value
            tick_diff = tick_close - tick_open
            tick_perc = tick_diff / tick_open
            
            On Error Resume Next
            
            Range("I" & sum_row).Value = tick_name
            Range("J" & sum_row).Value = tick_diff
            Range("K" & sum_row).Value = tick_perc
            Range("L" & sum_row).Value = tick_total
            sum_row = sum_row + 1

            tick_open = 0
            tick_close = 0
            tick_total = 0

        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
            tick_open = Cells(i, 3).Value
            tick_total = tick_total + Cells(i, 7).Value
        
        Else

            tick_total = tick_total + Cells(i, 7).Value
 
        End If

    Next i

    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'red/neg.conditional
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Color = 13551615
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'green/pos.conditional
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Color = 13561798
    End With
    Selection.FormatConditions(1).StopIfTrue = False


    'formatting
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Change %"
    Range("L1").Value = "Total Vol."

    Range("I1:L1").Select
    Selection.Font.Bold = True
    Selection.Interior.Color = 65535
    Selection.HorizontalAlignment = xlCenter
    
    Columns("J:J").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
    Columns("K:K").Select
    Selection.NumberFormat = "#,##0.0%;[Red](#,##0.0%)"
    
    Columns("L:L").Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
    
    Range("A1").Select

End Sub


