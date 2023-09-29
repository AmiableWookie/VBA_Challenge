Attribute VB_Name = "Module1"
Option Explicit

Sub CountStocks()

Dim LastRow As Long
Dim WorkSheetName, TickerName, GreatestPercentIncreaseTicker, GreatestPercentDecreaseTicker, GreatestVolumeTicker As String
Dim ws As Object
Dim i, row As Integer
Dim TickerChange, TickerChangePercent As Double
Dim YearOpenAlreadyCaptured As Boolean
Dim YearOpen, YearClose As Double
Dim percentchange As Double
Dim GreatestPercentIncrease, GreatestPercentDecrease As Double
Dim GreatestVolume, TotalVolume As LongLong



For Each ws In Worksheets

    row = 1
    YearOpenAlreadyCaptured = False
    YearOpen = 0
    YearClose = 0
    TickerChange = 0
    percentchange = 0
    TotalVolume = 0
    GreatestPercentDecrease = 0
    GreatestPercentIncrease = 0
    GreatestVolume = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'MsgBox (LastRow)
    
             
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Trade Volume"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        
    For i = 2 To LastRow
            
            'Setting opening value
            If YearOpenAlreadyCaptured = False Then
            
                YearOpen = ws.Cells(i, 3).Value
                
                YearOpenAlreadyCaptured = True
            
            End If
            
                    
        'Cut-off for one ticker changing to another
                 
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                YearClose = ws.Cells(i, 6).Value
                TickerChange = YearOpen - YearClose
                TickerName = ws.Cells(i, 1).Value
                
                'calculating percent change for each ticker
                
                If TickerChange = 0 Or YearOpen = 0 Then
                    ws.Cells(row, 11).Value = 0
                
                Else
                    percentchange = ((YearOpen - YearClose) / YearOpen)
                
                End If
                
                'calculating total volume for ticker
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                'moving row for calculated values input
                row = row + 1
            
                ws.Range("I" & row) = TickerName
                ws.Range("J" & row).Value = TickerChange
                ws.Range("J" & row).NumberFormat = "$0.00"
                ws.Range("K" & row).Value = percentchange
                ws.Range("K" & row).NumberFormat = "0.00%"
                ws.Range("L" & row).Value = TotalVolume
                ws.Range("L" & row).Style = "Normal"
                
                'coloring percent change column
                    If ws.Range("J" & row).Value >= 0 Then
                        ws.Range("J" & row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & row).Interior.ColorIndex = 3
                    End If
                
                ' Find the values for greatest decrease/increase and greatest volume.
                    If ws.Cells(row, 11).Value > GreatestPercentIncrease Then
                        GreatestPercentIncrease = ws.Cells(row, 11).Value
                        GreatestPercentIncreaseTicker = ws.Cells(row, 9).Value
                        
                    ElseIf ws.Cells(row, 11).Value < GreatestPercentDecrease Then
                        GreatestPercentDecrease = ws.Cells(row, 11).Value
                        GreatestPercentDecreaseTicker = ws.Cells(row, 9).Value
                        
                    ElseIf ws.Cells(row, 12).Value > GreatestVolume Then
                        GreatestVolume = ws.Cells(row, 12).Value
                        GreatestVolumeTicker = ws.Cells(row, 9).Value
                    End If
                
                TickerChange = 0
                YearOpenAlreadyCaptured = False
                YearClose = 0
                TotalVolume = 0
                
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
            
                       
    Next i


    If ws.Cells(row, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(row, 12).Value
        GreatestVolumeTicker = ws.Cells(row, 9).Value
    
    End If
    
'Setting values for percent increase, decrease, greatest volume, and associated tickers
ws.Cells(2, 16).Value = GreatestPercentIncrease
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = GreatestPercentDecrease
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = GreatestVolume

ws.Cells(2, 15).Value = GreatestPercentIncreaseTicker
ws.Cells(3, 15).Value = GreatestPercentDecreaseTicker
ws.Cells(4, 15).Value = GreatestVolumeTicker


Next ws

End Sub
