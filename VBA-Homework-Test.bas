Attribute VB_Name = "Module1"
Sub StockNumbers()

Dim ws As Worksheet

For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim YearlyChange As Double
YearlyChange = 0
Dim PercentChange As Double
PercentChange = 0
Dim TotalVolume As LongLong
TotalVolume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Opening_Price As Double
Opening_Price = 0
Dim Closing_Price As Double
Closing_Price = 0


ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

For i = 2 To LastRow

    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        Opening_Price = ws.Cells(i, 3).Value
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i + 1, 7).Value
        Closing_Price = ws.Cells(i, 6).Value
        YearlyChange = Closing_Price - Opening_Price
        PercentChange = (YearlyChange / Opening_Price) * 100
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        ws.Range("K" & Summary_Table_Row).Value = PercentChange
        ws.Range("L" & Summary_Table_Row).Value = TotalVolume
        Summary_Table_Row = Summary_Table_Row + 1
        TotalVolume = 0
        YearlyChange = 0
        PercentChange = 0
    
    Else
         TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
    End If
    
                If Cells(i, 10) > 0 Then
                    Cells(i, 10).Interior.Color = vbGreen
    
                 ElseIf Cells(i, 10) < 0 Then
                    Cells(i, 10).Interior.Color = vbRed
                End If
    

Next i

Next ws

End Sub
