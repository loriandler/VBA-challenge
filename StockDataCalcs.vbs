Attribute VB_Name = "RunCalcs"
Sub OnAllSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call StockDataChallenge
    Next
    Application.ScreenUpdating = True
End Sub
Sub StockDataChallenge()


' Define all of the variables

Dim ticker As String
Dim lastRow As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double
Dim summaryTableRow As Long
Dim rng As Range


' Find last row of data
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


' Add headers to worksheet
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"


' Set summary table row
summaryTableRow = 2

' Loop through data
For I = 2 To lastRow

    ' Find a Change in ticker symbol and set values
    If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        
        ticker = Cells(I, 1).Value
        
        openingPrice = Cells(I, 3).Value
        
        closingPrice = Cells(I, 6).Value
        
        yearlyChange = closingPrice - openingPrice
        
        ' Calculate the Percent Change
        If openingPrice <> 0 Then
            percentChange = yearlyChange / openingPrice
            
        Else
            percentChange = 0
    
        End If
        
        ' Add up the total volume
        totalVolume = WorksheetFunction.Sum(Range(Cells(I, 7), Cells(I + 1, 7)))
        
        ' Display results
        Range("I" & summaryTableRow).Value = ticker
        Range("J" & summaryTableRow).Value = yearlyChange
        Range("K" & summaryTableRow).Value = percentChange
        Range("L" & summaryTableRow).Value = totalVolume
        
        ' format percentage
        Range("K" & summaryTableRow).NumberFormat = "0.00%"
        Range("J" & summaryTableRow).NumberFormat = "$0.00"
        
         If yearlyChange < 0 Then
        Range("J" & summaryTableRow).Interior.ColorIndex = 3
    Else
        Range("J" & summaryTableRow).Interior.ColorIndex = 4
        
        
    End If
        ' next row
        summaryTableRow = summaryTableRow + 1
        
    End If
Next I

    ' Search and display Greatest Values in summary table
    Dim MaxIncrease As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecrease As Double
    Dim MaxDecreaseTicker As String
    Dim GreatestTotalVolume As Double, x As Long
    Dim GreatestTotalVolumeTicker As String
    
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
    
     ' Format percentage
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' format the column size
    Range("I1:Q1").EntireColumn.AutoFit
    

End Sub
    
