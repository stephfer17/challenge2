Attribute VB_Name = "Module2"
Sub yearly_change()
'Ask BSC helped me
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
'Trying to get the info from one column and bring it to another
    lastrow = ws.Cells(Rows.Count, 2).End(xlUp).Row
'For the yearly change making eaiser to understand
    Dim yearlast As Double
    Dim yearfirst As Double
    Dim change As Double
    Dim percentage As Integer
    j = 2
    ' Starting the loop
    For i = 2 To lastrow
'definingthe change

        If yearfirst = 0 Then
        yearfirst = ws.Cells(i, 3).Value
        End If
' changing the if the statement making the logic fixing to make more sense if and else and if
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
      yearlast = ws.Cells(i, 6).Value
        change = yearlast - yearfirst
'defining percentage
        Percent = change / yearfirst * 100
'putting the in the right cells
        ws.Cells(j, 10).Value = change
        ws.Cells(j, 11).Value = Percent
        j = j + 1
 ' Conditionals

     If ws.Cells(j, 10).Value > 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
        Else: ws.Cells(j, 10).Interior.ColorIndex = 0
    
    End If
    If ws.Cells(j, 11).Value > 0 Then
     ws.Cells(j, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 11).Value < 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 3
        Else: ws.Cells(j, 11).Interior.ColorIndex = 0
        End If
	End If
    Next i
       ws.Range("P4") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
    ws.Range("P5") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
'WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" &lastrow)), Range("K2:K" & rowCount), 0)
'Same thing with the ticker module but its max and min for the least 
    Next ws
End Sub


