VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-------------------
'VBA script begins here:
'-------------------

'This sub allows the code in the 'stocks' to be applied to each worksheet once run.
Sub sheetloop()

'Creating the ws variable and declaring it as a worksheet.
Dim ws As Worksheet
    
    'This will loop through each worksheet(declared as ws above) in the workbook and call the below sub 'stocks' for each of them.
    For Each ws In ThisWorkbook.Sheets
        Call stocks(ws)
    Next ws
    
End Sub

'New sub to organize and analyze the data
Sub stocks(ws As Worksheet)

'Variable to identify the last rows in the worksheet and the
Dim lastrow As Double

'Our iteration variables
Dim i As Double
Dim j As Double

'Used to identify the yearly change and percent change for each ticker
Dim TickerVariable As Double
Dim open_value As Double
Dim close_value As Double
Dim Volume_Counter As Double

'These are used to output the greatest volume and percent change at the end of the sub
Dim GreatestTickerIncrease As String
Dim GreatestTickerDecrease As String
Dim GreatestTickerVolume As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

'Obtain the last row in the sheet and initialize the TickerVariable to 1
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TickerVariable = 1

'Initialize the greatest increase and greatest decrease variables
GreatestIncrease = ws.Cells(2, 11).Value
GreatestDecrease = ws.Cells(2, 11).Value
GreatestVolume = ws.Cells(2, 12).Value

'Input the column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % increase"
ws.Cells(3, 14).Value = "Greatest % decrease"
ws.Cells(4, 14).Value = "Greatest total volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'This loop aims to start on the second row until the last row in the sheet to get the ticker, difference between.\
'the opening and closing stock prices, and the volume.
For i = 2 To lastrow

    'Outputing the difference of the first row's ticker when the TickerVariable = 1
    If TickerVariable = 1 Then
    
        TickerVariable = TickerVariable + 1
        open_value = ws.Cells(i + 1, 3).Value
        
    Else
        'Returning the difference between the opening and closing prices for each unique ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Return the ticker
            ws.Cells(TickerVariable, 9).Value = ws.Cells(i, 1).Value
            
            'Put the closing value into a variable called "close_value" and take subtract it from the opening price
            close_value = ws.Cells(i, 6).Value
            ws.Cells(TickerVariable, 10).Value = close_value - open_value
        
            'By making this IF statement, we are preventing a DIV/0 error
            If open_value <> 0 Then
                
                'Divide the yearly change with the opening value for each stock ticker
                ws.Cells(TickerVariable, 11).Value = ws.Cells(TickerVariable, 10).Value / open_value
            Else
                
                'The percent change will return 0 if the opening value price is 0 to avoid a divide by 0 error
                ws.Cells(TickerVariable, 11).Value = 0
            
            End If
        
            'This will sum the total volume for each stock ticker and return it with the corresponding ticker
            Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value
            ws.Cells(TickerVariable, 12).Value = Volume_Counter
        
            'Create the opening value for each stock ticker and reset the ticker counter and volume counter variables
            open_value = ws.Cells(i + 1, 3).Value
            TickerVariable = TickerVariable + 1
            Volume_Counter = 0
                
        Else
        
            Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value

        End If
    
    End If
    
Next i

'This will loop through the data created in the first loop to retrieve the ticker with the.\
'greatest percent increase, greatest decrease, and highest volume
For j = 2 To 3001
    
    'Returns the ticker with the greatest percent increase
    If ws.Cells(j, 11).Value > GreatestIncrease Then
        
        GreatestIncrease = ws.Cells(j, 11).Value
        GreatestTickerIncrease = ws.Cells(j, 9).Value
        
    End If
        
    'Returns the ticker with the greatest percent decrease
    If ws.Cells(j, 11).Value < GreatestDecrease Then
    
        GreatestDecrease = ws.Cells(j, 11).Value
        GreatestTickerDecrease = ws.Cells(j, 9).Value
        
    End If

    'Returns the ticker with the highest volume
    If ws.Cells(j, 12).Value > GreatestVolume Then
        
        GreatestVolume = ws.Cells(j, 12).Value
        GreatestTickerVolume = ws.Cells(j, 9).Value
        
        
    End If
    
    'This will format the cells for column J—green for increase in stock price and red for decrease in stock price
    If ws.Cells(j, 10).Value < 0 Then
    
        ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            
    Else
        ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            
    End If
    

Next j

'Output the greatest increase, decrease, and highest volume into their corresponding cells
ws.Cells(2, 15).Value = GreatestTickerIncrease
ws.Cells(3, 15).Value = GreatestTickerDecrease
ws.Cells(4, 15).Value = GreatestTickerVolume

ws.Cells(2, 16).Value = GreatestIncrease
ws.Cells(3, 16).Value = GreatestDecrease
ws.Cells(4, 16).Value = GreatestVolume

End Sub
