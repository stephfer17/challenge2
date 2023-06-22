Attribute VB_Name = "Module1"
Sub tickers()
' Ask BSC helped me
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
' Just setting all the nescarry colums and rows for the asigment
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change$"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total stock volume"
' Setting up places names
    ws.Range("O2").Value = "ticker"
    ws.Range("P2").Value = "value"
    ws.Cells(3, 14).Value = "Max Volume"
    ws.Cells(4, 14).Value = "Greatest increase %"
    ws.Cells(5, 14).Value = "Greastest decrese %"
'Setting the varaible as ticker
    Dim ticker As String
    Dim volume As Double
    Dim total_volume As Double
    Dim maxvol As Double
'Trying to get the info from one column and bring it to another
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' AskBSC helped with this part
      j = 2
'The loop
    For i = 2 To lastrow
    ticker = ws.Cells(i, 1).Value
' Checking for the ticker is the same or different
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    ' the next two lines were helped but AskBSC
    'printing out the values
     ws.Cells(j, 9).Value = ticker
     j = j + 1
    Else
'add the spefic cells for volume
    volume = ws.Cells(i, 7).Value
    total_volume = volume + ws.Cells(i, 7).Value
     'Print the volume
    ws.Cells(j, 12).Value = total_volume
    End If
' For finding max volume orginal idea
     '  If Cells(i, 7).Value > Cells(i + 1, 7).Value Then
      ' maxvol = Cells(i, 7).Value
      ' MsgBox (maxvol)
      ' End IF

   Next i
   ' Ask BSC helped
   ws.Range("P3") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
' I know this is the code but it not working
' WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" &lastrow)), Range("L2:L" & rowCount), 0)



   Next ws
End Sub


