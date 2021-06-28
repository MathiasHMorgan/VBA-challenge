# VBA-challenge
Homework 2

Sub StocksData():

'Dimesion Variables

Dim Ticker_Name As String
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim Volume As Double

Dim ws As Worksheet

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double

Dim Results_Table_Row As Integer

Results_Table_Row = 2

'loop through all worksheets

For Each ws In Worksheets
   
'Identify Last Row

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Headers required

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    

'Loop Through Data and extract values

    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Tickers
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("I" & Results_Table_Row).Value = Ticker_Name
        
        'Yearly Change
        Year_Open = ws.Cells(i, 3).Value
        Year_Close = ws.Cells(i, 6).Value
        
        Yearly_Change = Year_Close - Year_Open
        ws.Range("J" & Results_Table_Row).Value = Yearly_Change

        'Percent Change
        Percentage_Change = (Year_Close / Year_Open) - 1
        ws.Range("K" & Results_Table_Row).Value = Percentage_Change
        ws.Range("K" & Results_Table_Row).NumberFormat = "0.00%"
        
        'Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        ws.Range("L" & Results_Table_Row).Value = Total_Stock_Volume
         
        'Loop Through all required cells
        
         Results_Table_Row = Results_Table_Row + 1
         
       End If
       
       'Color Index
        If ws.Range("J" & Results_Table_Row).Value <= 0 Then
        ws.Range("J" & Results_Table_Row).Interior.ColorIndex = 3
        
        ElseIf ws.Range("J" & Results_Table_Row).Value > 0.001 Then
        ws.Range("J" & Results_Table_Row).Interior.ColorIndex = 4
        
        
        End If
        
        'Find Values (Greateste Increase, Greatest Decrease & Greatest Total Volume)
        ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCount))
        ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCount))
        ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
    
        'Find Corresponding Ticker Values
        GreatestIncrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("O2") = Cells(GreatestIncrease + 1, 9)
        
        GreatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("O3") = Cells(GreatestDecrease + 1, 9)
        
        GreatestTotalVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("O4") = Cells(GreatestTotalVolume + 1, 9)

    Next i
      
Next ws
    
End Sub
