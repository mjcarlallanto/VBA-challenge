Attribute VB_Name = "Module1"

===========================================================

Sub runAll()

Call summary
Call stocks
Call bonus

End Sub

===========================================================

Sub summary()

  Dim ws As Worksheet
  
  For Each ws In ThisWorkbook.Worksheets
  
    Dim ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim start_date As Double
    Dim end_date As Double
    
    Dim total As Double
    total = 0
    
    Dim summary_row As Integer
    summary_row = 2
  
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ws.Range("J1").value = "Ticker"
    ws.Range("K1").value = "Opening Price"
    ws.Range("L1").value = "Closing Price"
    ws.Range("M1").value = "Start Date"
    ws.Range("N1").value = "Close Date"
    ws.Range("O1").value = "Total Stock Volume"
  
    opening_price = ws.Cells(2, 3).value
    start_date = ws.Cells(2, 2).value
    
    ws.Range("K2").value = opening_price
    ws.Range("M2").value = start_date
  
    For I = 2 To lastrow

        If ws.Cells(I + 1, 1).value <> ws.Cells(I, 1).value Or Left(ws.Cells(I + 1, 2).value, 4) <> Left(ws.Cells(I, 2).value, 4) Then

            ticker = ws.Cells(I, 1).value
            closing_price = ws.Cells(I, 6).value
            end_date = ws.Cells(I, 2).value
            opening_price = ws.Cells(I + 1, 3).value
            start_date = ws.Cells(I + 1, 2).value
            total = total + ws.Cells(I, 7).value

            ws.Range("J" & summary_row).value = ticker
            ws.Range("L" & summary_row).value = closing_price
            ws.Range("N" & summary_row).value = end_date
            ws.Range("O" & summary_row).value = total

            summary_row = summary_row + 1
            ws.Range("K" & summary_row).value = opening_price
            ws.Range("M" & summary_row).value = start_date
      
            total = 0
    
    
        Else

            total = total + ws.Cells(I, 7).value

            ws.Range("O" & summary_row).value = total
        
        End If

    Next I
Next ws
End Sub

===========================================================

Sub stocks()

Dim ws As Worksheet
  
For Each ws In ThisWorkbook.Worksheets

    Dim yearly_change As Double
    Dim percent_change As Double
    Dim summary_row As Integer
    summary_row = 2
  
    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    ws.Range("R1").value = "Ticker"
    ws.Range("S1").value = "Yearly Change"
    ws.Range("T1").value = "Percent Change"
    ws.Range("U1").value = "Total Stock Volume"

    For I = 2 To lastrow
        yearly_change = ws.Cells(I, 12).value - ws.Cells(I, 11).value
        If yearly_change > 0 Then
            ws.Range("S" & summary_row).Interior.ColorIndex = 4
        Else
            ws.Range("S" & summary_row).Interior.ColorIndex = 3
        End If
        If ws.Cells(I, 11).value = 0 Then
            percent_change = 0
        Else
             percent_change = yearly_change / ws.Cells(I, 11).value
        End If
        
        ws.Range("R" & summary_row).value = ws.Cells(I, 10).value
        ws.Range("S" & summary_row).value = yearly_change
        ws.Range("T" & summary_row).value = FormatPercent(percent_change)
        ws.Range("U" & summary_row).value = ws.Cells(I, 15).value
    
        summary_row = summary_row + 1
    Next I
Next ws
End Sub

===========================================================

Sub bonus()

Dim ws As Worksheet
  
For Each ws In ThisWorkbook.Worksheets

    Dim ticker As String
    Dim value As Double
  
    lastrow = ws.Cells(Rows.Count, 18).End(xlUp).Row

    ws.Range("X2").value = "Greatest % Increase"
    ws.Range("X3").value = "Greatest % Decrease"
    ws.Range("X4").value = "Greatest Total Volume"
    ws.Range("Y1").value = "Ticker"
    ws.Range("Z1").value = "Value"
    ws.Range("Y2").value = ws.Range("R2").value
    ws.Range("Y3").value = ws.Range("R2").value
    ws.Range("Y4").value = ws.Range("R2").value
    ws.Range("Z2").value = ws.Range("T2").value
    ws.Range("Z3").value = ws.Range("T2").value
    ws.Range("Z4").value = ws.Range("U2").value

    For I = 2 To lastrow
        If ws.Range("Z2").value < ws.Cells(I, 20).value Then
            ws.Range("Z2").value = FormatPercent(ws.Cells(I, 20).value)
            ws.Range("Y2").value = ws.Cells(I, 18).value
        End If
        If ws.Range("Z3").value > ws.Cells(I, 20).value Then
            ws.Range("Z3").value = FormatPercent(ws.Cells(I, 20).value)
            ws.Range("Y3").value = ws.Cells(I, 18).value
        End If
        If ws.Range("Z4").value < ws.Cells(I, 21).value Then
            ws.Range("Z4").value = ws.Cells(I, 21).value
            ws.Range("Y4").value = ws.Cells(I, 18).value
        End If
    Next I

Next ws
End Sub
