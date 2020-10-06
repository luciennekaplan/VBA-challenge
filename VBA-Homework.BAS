Attribute VB_Name = "Module1"
 Sub Stocks()

Dim ws As Worksheet

For Each ws In Worksheets

Dim i As Long
Dim Ticker As String
Dim Current_Ticker As String
Dim Next_Ticker As String
Dim YearlyChange As Double
YearlyChange = 0
Dim PercentChange As Long
PercentChange = 0
Dim TotalVolume As LongLong
TotalVolume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Opening_Price As Double
Opening_Price = 0
Dim Closing_Price As Double
Closing_Price = 0
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

For i = 2 To LastRow
Current_Ticker = ws.Cells(i, 1).Value
Next_Ticker = ws.Cells(i + 1, 1)

    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        Opening_Price = ws.Cells(i, 3).Value
    
    ElseIf Next_Ticker <> Current_Ticker Then
        Current_Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i + 1, 7).Value
        Closing_Price = ws.Cells(i, 6).Value
        YearlyChange = Closing_Price - Opening_Price
            If Opening_Price = 0 Then
            PercentChange = 0
            ElseIf Opening_Price <> 0 Then
            PercentChange = (YearlyChange / Opening_Price) * 100
            End If
        ws.Range("I" & Summary_Table_Row).Value = Current_Ticker
        ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        ws.Range("K" & Summary_Table_Row).Value = PercentChange
        ws.Range("L" & Summary_Table_Row).Value = TotalVolume
        Summary_Table_Row = Summary_Table_Row + 1
        TotalVolume = 0
        YearlyChange = 0
        PercentChange = 0
    
    ElseIf Current_Ticker = Next_Ticker Then
         TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
    End If
    
    
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.Color = vbGreen
    
                 ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.Color = vbRed
                 End If
    

Next i

Next ws

End Sub

