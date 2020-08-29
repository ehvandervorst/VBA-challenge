Sub WallStreet()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets

Dim Ticker As String
Dim Volume As String
Volume = 0

Dim Open_Price As Double
Dim Close_Price As Double
Dim Price_Change As Double
Dim Percent_Change As Double
Dim close_value As Double
Dim open_value As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim LastRow As String

LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    Ticker = ws.Cells(i, 1).Value
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    Volume = Volume + ws.Cells(i, 7).Value
    ws.Range("L" & Summary_Table_Row).Value = Volume
    
    close_value = ws.Cells(i, 1).Offset(0, 5).Value
    
    open_value = ws.Cells(i - 261, 1).Offset(0, 2).Value
        
    Price_Change = close_value - open_value
    ws.Range("J" & Summary_Table_Row).Value = Price_Change
       
    Percent_Change = (Price_Change / open_value) * 100
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    
    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
Summary_Table_Row = Summary_Table_Row + 1
Volume = 0
Price_Change = 0
Percent_Change = 0

Else
Volume = Volume + ws.Cells(i, 7).Value

End If

Next i
    
Next ws

End Sub
