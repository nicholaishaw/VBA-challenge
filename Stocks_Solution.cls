VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub allsheets()
Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        Call stocks(ws)
    Next ws
    
End Sub

Sub stocks(ws As Worksheet)
Dim lastrow As Double
Dim FinalLastRow As Double
Dim i As Double
Dim j As Double
Dim GreatestTickerIncrease As String
Dim GreatestTickerDecrease As String
Dim GreatestTickerVolume As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim TickerVariable As Double
Dim open_value As Double
Dim close_value As Double
Dim Volume_Counter As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TickerVariable = 1

GreatestIncrease = ws.Cells(2, 11).Value
GreatestDecrease = ws.Cells(2, 11).Value
GreatestVolume = ws.Cells(2, 12).Value

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % increase"
ws.Cells(3, 14).Value = "Greatest % decrease"
ws.Cells(4, 14).Value = "Greatest total volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

For i = 2 To lastrow
    If TickerVariable = 1 Then
    
        TickerVariable = TickerVariable + 1
        open_value = ws.Cells(i + 1, 3).Value
        
    Else
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        
            ws.Cells(TickerVariable, 9).Value = ws.Cells(i, 1).Value
        
            close_value = ws.Cells(i, 6).Value
            ws.Cells(TickerVariable, 10).Value = close_value - open_value
        
            If open_value <> 0 Then
        
                ws.Cells(TickerVariable, 11).Value = ws.Cells(TickerVariable, 10).Value / open_value
            Else
                ws.Cells(TickerVariable, 11).Value = 0
            
            End If
        
            Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value
            ws.Cells(TickerVariable, 12).Value = Volume_Counter
        
            open_value = ws.Cells(i + 1, 3).Value
            TickerVariable = TickerVariable + 1
            Volume_Counter = 0
                
        
        Else
            Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value

        End If
    
    End If
    
Next i


For j = 2 To 3001
    
    If ws.Cells(j, 11).Value > GreatestIncrease Then
        
        GreatestIncrease = ws.Cells(j, 11).Value
        GreatestTickerIncrease = ws.Cells(j, 9).Value
        
    End If
        
        
    If ws.Cells(j, 11).Value < GreatestDecrease Then
    
        GreatestDecrease = ws.Cells(j, 11).Value
        GreatestTickerDecrease = ws.Cells(j, 9).Value
        
    End If

    If ws.Cells(j, 12).Value > GreatestVolume Then
        
        GreatestVolume = ws.Cells(j, 12).Value
        GreatestTickerVolume = ws.Cells(j, 9).Value
        
        
    End If
    
    
    If ws.Cells(j, 10).Value < 0 Then
    
        ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            
    Else
        ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            
    End If
    

Next j

ws.Cells(2, 15).Value = GreatestTickerIncrease
ws.Cells(3, 15).Value = GreatestTickerDecrease
ws.Cells(4, 15).Value = GreatestTickerVolume

ws.Cells(2, 16).Value = GreatestIncrease
ws.Cells(3, 16).Value = GreatestDecrease
ws.Cells(4, 16).Value = GreatestVolume




End Sub
