Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        summaryRow = 2
        lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        Dim i As Long
        i = 2
        
        Do While i <= lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            volume = 0
            
            Do While ws.Cells(i, 1).Value = ticker
                volume = volume + ws.Cells(i, 7).Value
                closePrice = ws.Cells(i, 6).Value
                i = i + 1
            Loop
            
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If
            
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = quarterlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = volume
            
           
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                increaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                decreaseTicker = ticker
            End If
            
            If volume > greatestVolume Then
                greatestVolume = volume
                volumeTicker = ticker
            End If
            
            summaryRow = summaryRow + 1
        Loop
        
        ' Output the results
        ws.Cells(2, 14).Value = "Greatest % Increase: " & increaseTicker & " (" & greatestIncrease & "%)"
        ws.Cells(3, 14).Value = "Greatest % Decrease: " & decreaseTicker & " (" & greatestDecrease & "%)"
        ws.Cells(4, 14).Value = "Greatest Total Volume: " & volumeTicker & " (" & greatestVolume & ")"
    Next ws
End Sub

