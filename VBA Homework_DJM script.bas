Attribute VB_Name = "Module1"
Sub AlphaTesting()
For Each ws In Worksheets
    
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10) = "Ticker Vol"
    
    Dim Ticker As String
    Dim TickerVol As Double
    TickerVol = 0
    
    Dim Tablestart As Integer
    Tablestart = 2
    
    
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TickerVol = TickerVol + ws.Cells(i, 7).Value
        ws.Range("I" & Tablestart).Value = Ticker
        ws.Range("J" & Tablestart).Value = TickerVol
        Tablestart = Tablestart + 1
        TickerVol = 0
    Else
        TickerVol = TickerVol + ws.Cells(i, 7).Value
    End If
    
    Next i
    
 Next ws
    
End Sub

