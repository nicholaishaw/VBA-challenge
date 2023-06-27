VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()
Dim lastrow As Long
Dim FinalLastRow As Integer
Dim i As Long
Dim j As Integer
Dim TickerVariable As Integer
Dim open_value As Integer
Dim close_value As Integer
Dim Volume_Counter As Long


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
TickerVariable = 1

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % increase"
Cells(3, 14).Value = "Greatest % decrease"
Cells(4, 14).Value = "Greatest total volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

For i = 2 To lastrow
        
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        TickerVariable = TickerVariable + 1
        Cells(TickerVariable, 9).Value = Cells(i, 1).Value
        
        open_value = Cells(i, 3).Value
        close_value = Cells(i, 6).Value
        Cells(TickerVariable, 10).Value = open_value - close_value
        
        Cells(i, 11).Value = Cells(i, 10).Value / (open_value + close_value)
    
        'Volume_Counter = Volume_Counter + Cells(i, 7).Value
        'Cells(i, 12).Value = Volume_Counter
        
        open_value = 0
        close_value = 0
        Volume_Counter = 0
        
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
    End If
    
Next i

For j = 2 To FinalLastRow
    
    If Cells(j, 11).Value > Cells(j + 1, 11).Value Then
        
        Cells(2, 15).Value = Cells(j, 9).Value
        Cells(2, 16).Value = Cells(j, 11).Value
        
    ElseIf Cells(j, 11).Value < Cells(j + 1, 11).Value Then
    
        Cells(3, 15).Value = Cells(j, 9).Value
        Cells(3, 16).Value = Cells(j, 11).Value
        
    End If

    If Cells(j, 12).Value > Cells(j + 1, 12).Value Then
        
        Cells(4, 15).Value = Cells(j, 9).Value
        Cells(4, 16).Value = Cells(j, 12).Value
    End If

Next j


End Sub

