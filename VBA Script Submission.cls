VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData()

' Variable definition key
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    Dim LastRow As Long
    Dim r As Long
    Dim SummaryTableRow As Integer
        SummaryTableRow = 2
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentageChange As Double
    Dim firstrow As Boolean
    Dim TotalVolume As Double

   firstrow = True
    
    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "QuarterlyChange"
    ws.Cells(1, 11).Value = "PercentChange"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    ' Find the last row of the sheet
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows of data
    For r = 2 To LastRow
    
    'Set OpeningPrice
    OpeningPrice = ws.Cells(r, 3).Value
     
        ' Check if we are at the end of a ticker or the last row
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Or r = LastRow Then

            ' Ticker symbol
            Ticker = ws.Cells(r, 1).Value
            ws.Cells(SummaryTableRow, 9).Value = Ticker
            
            ' Total Volume
            TotalVolume = TotalVolume + Cells(r, 7).Value
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume

            ' ClosingPrice
            ClosingPrice = ws.Cells(r, 6).Value

            ' Calculate the quarterly change
            QuarterlyChange = ClosingPrice - OpeningPrice
            ws.Cells(SummaryTableRow, 10).Value = QuarterlyChange
            
            'Percentage Change
            PercentageChange = (QuarterlyChange / OpeningPrice) * 100
            ws.Cells(SummaryTableRow, 11).Value = Format(PercentageChange, "0.00") & "%"
            
            'reset opening price
            OpeningPrice = Cells(r, 3)
            TotalVolume = 0

            ' Move to the next summary row
            SummaryTableRow = SummaryTableRow + 1

         
        Else
            ' This is the first row for a new ticker, so store the opening price
            ' OpeningPrice = ws.Cells(r, 3).Value
            TotalVolume = TotalVolume + Cells(r, 7).Value
            firstrow = False
            
            
        End If
        
            'Cell Colors
            If (ws.Cells(r, 10).Value > 0 Or ws.Cells(r, 10).Value = 0) Then
                 ws.Cells(r, 10).Interior.ColorIndex = 10
            ElseIf (ws.Cells(r, 10).Value < 0) Then
                 ws.Cells(r, 10).Interior.ColorIndex = 3
            End If
        
    Next r

    ' New loop for new table
             For k = 2 To LastRow
             lastRow2 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
             
    ' Find Greatest % Increase, Greatest % Decrease, Greatest Total Volume
       If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow2)) Then
        ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
        ws.Cells(2, 16).Value = ws.Cells(k, 11).Value
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow2)) Then
        ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
        ws.Cells(3, 16).Value = ws.Cells(k, 11).Value
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow2)) Then
        ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
        ws.Cells(4, 16).Value = ws.Cells(k, 12).Value
        End If
    Next k

Next ws

End Sub
