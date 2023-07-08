Attribute VB_Name = "Reset"
Sub OnAllSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call ResetSheets
    Next
    Application.ScreenUpdating = True
End Sub

Sub ResetSheets()
    Range("I1:Q1").EntireColumn.Clear
End Sub
