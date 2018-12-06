Attribute VB_Name = "Module1"
Sub TICKER()
    Dim ws As Worksheet
    
    Dim TickerName As String
    
    Dim TickerVolume As Double
    TickerVolume = 0
    
    Dim TickerRow As Integer
    TickerRow = 2
    
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For Each ws In Worksheets
            
            
                For i = 2 To LastRow
                
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        TickerName = ws.Cells(i, 1).Value
                        TickerVolume = TickerVolume + ws.Cells(i, 7)
                        ws.Range("J" & TickerRow).Value = TickerName
                        ws.Range("K" & TickerRow).Value = TickerVolume
                        TickerRow = TickerRow + 1
                        TickerVolume = 0
                        
                    Else
                        TickerVolume = TickerVolume + ws.Cells(i, 7).Value
                    
                    End If
                    
                Next i
                
    Next ws
End Sub
