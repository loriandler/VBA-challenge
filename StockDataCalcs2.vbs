Attribute VB_Name = "RunCalcs"
Sub StockDataChallenge():

' Set variable for Worksheets
Dim ws As Worksheet

' Begin loop for all worksheets
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

    ' Define all variables
    Dim Ticker As String
    Dim i As Long
    Dim totalVolume As Double
    Dim yearlyChange As Double
    Dim openingPrice As Double
    Dim Percent_Change As Double
    Dim summaryTableRow As Integer
    
    
    
    ' Define opening values
    totalVolume = 0
    yearlyChange = 0
    openingPrice = Cells(2, 3).Value
    Percent_Change = 0
    summaryTableRow = 2
    
    ' Determine Last Row in Complete List of Tickers
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Loop through data
        For i = 2 To lastRow
    
            ' Check to make sure still within same Ticker, if not
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
                ' Set Formulas
                Ticker = Cells(i, 1).Value
                
                yearlyChange = Cells(i, 6).Value - openingPrice
            
                Percent_Change = (yearlyChange / openingPrice)
                
                openingPrice = Cells(i + 1, 3).Value
                
                totalVolume = totalVolume + Cells(i, 7).Value
        
                ' Print data in summary table
                Range("I" & summaryTableRow).Value = Ticker
                
                Range("J" & summaryTableRow).Value = yearlyChange
                
                Range("K" & summaryTableRow).Value = Format(Percent_Change, "0.00%")
                
                Range("L" & summaryTableRow).Value = totalVolume
                
                ' Color Positive and Negative Yearly Change
                If Range("J" & summaryTableRow).Value < 0 Then
                    
                    Range("J" & summaryTableRow).Interior.ColorIndex = 3
                        
                End If
                
                If Range("J" & summaryTableRow).Value > 0 Then
                    
                    Range("J" & summaryTableRow).Interior.ColorIndex = 4
                
                End If
                
                
        
                summaryTableRow = summaryTableRow + 1
            
                totalVolume = 0
                
            Else
            
                totalVolume = totalVolume + Cells(i, 7).Value
                
            End If
            
        Next i
        
    ' Add headers to columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

       ' Search and display Greatest Values in summary table
    Dim MaxIncrease As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecrease As Double
    Dim MaxDecreaseTicker As String
    Dim GreatestTotalVolume As Double, x As Long
    Dim GreatestTotalVolumeTicker As String
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    
    For x = 2 To lastRow
        If Range("K" & x).Value >= MaxIncrease Then
            MaxIncrease = Range("K" & x).Value
            MaxIncreaseTicker = Range("I" & x).Value
            
        End If
    Next x
    Range("Q2").Value = MaxIncrease
    Range("P2").Value = MaxIncreaseTicker

    
    For x = 2 To lastRow
        If Range("K" & x).Value <= MaxDecrease Then
            MaxDecrease = Range("K" & x).Value
            MaxDecreaseTicker = Range("I" & x).Value
        End If
    Next x
    
    Range("Q3").Value = MaxDecrease
    Range("P3").Value = MaxDecreaseTicker
    
    
    For x = 2 To lastRow
        If Range("L" & x).Value >= GreatestTotalVolume Then
            GreatestTotalVolume = Range("L" & x).Value
            GreatestTotalVolumeTicker = Range("I" & x).Value
            
        End If
    Next x
    
    Range("Q4").Value = GreatestTotalVolume
    Range("P4").Value = GreatestTotalVolumeTicker
    

    
    ' Autofit columns
        ws.Columns("A:Q").AutoFit
    
    
Next ws
End Sub

