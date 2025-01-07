Sub QuarterlyStockChanges()

    Dim ws As Worksheet
    
    ' define variables for maxes
    Dim MaxInc As Double
    Dim MaxDec As Double
    Dim MaxVol As Double
    Dim TickerMaxInc As String
    Dim TickerMaxDec As String
    Dim TickerMaxVol As String
    
    MaxInc = 0
    MaxDec = 0
    MaxVol = 0
    
    ' define variables for columns
    Dim OpenCol As Integer
    Dim CloseCol As Integer
    Dim VolCol As Integer
    
    OpenCol = 3 ' <open>
    CloseCol = 6 ' <close>
    VolCol = 7 ' <vol>
    
    ' loop through all sheets
        For Each ws In ThisWorkbook.Worksheets
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        Dim TotalVol As Double
        Dim StartRow As Long
        
        ' assign variable
        TotalVol = 0
        
        ' column headers in row 1 so start at row 2
        StartRow = 2
        
        ' add output headers to right of data
        With ws
            ' column I
            .Cells(1, 9).Value = "Ticker"
            ' column J
            .Cells(1, 10).Value = "Quarterly Change"
            ' column K
            .Cells(1, 11).Value = "Percent Change"
            ' column L
            .Cells(1, 12).Value = "Total Stock Volume"
        End With
        
        Dim OutputRow As Long
        
        ' results will start writing at this row
        OutputRow = 2
        
        ' loop through rows to calc metrics
        Dim i As Long
    For i = 2 To LastRow
            If Not IsNumeric(ws.Cells(i, OpenCol).Value) Or IsEmpty(ws.Cells(i, OpenCol).Value) Then ws.Cells(i, OpenCol).Value = 0
            If Not IsNumeric(ws.Cells(i, CloseCol).Value) Or IsEmpty(ws.Cells(i, CloseCol).Value) Then ws.Cells(i, CloseCol).Value = 0
            If Not IsNumeric(ws.Cells(i, VolCol).Value) Or IsEmpty(ws.Cells(i, VolCol).Value) Then ws.Cells(i, VolCol).Value = 0
            
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Dim Ticker As String
                Dim OpenPrice As Double
                Dim ClosePrice As Double
                Ticker = ws.Cells(i, 1).Value
                
                ' <open> column
                OpenPrice = ws.Cells(StartRow, OpenCol).Value
                
                ' <close> column
                ClosePrice = ws.Cells(i, CloseCol).Value
                
                ' <vol> column
                TotalVol = TotalVol + ws.Cells(i, VolCol).Value
                
                Dim PercentChange As Double
                
                ' calc percent change
                If OpenPrice <> 0 Then
                    PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                Else
                    PercentChange = 0
                End If
                
                ' show output data to the same sheet data came from
                With ws
                    .Cells(OutputRow, 9).Value = Ticker
                    .Cells(OutputRow, 10).Value = ClosePrice - OpenPrice ' the quarterly change
                    .Cells(OutputRow, 11).Value = Format(PercentChange, "0.00%")
                    .Cells(OutputRow, 12).Value = TotalVol
                End With
                
                ' update min/max
                If PercentChange > MaxInc Then
                    MaxInc = PercentChange
                    TickerMaxInc = Ticker
                End If
                If PercentChange < MaxDec Then
                    MaxDec = PercentChange
                    TickerMaxDec = Ticker
                End If
                If TotalVol > MaxVol Then
                    MaxVol = TotalVol
                    TickerMaxVol = Ticker
                End If
                
                ' reset it for the next ticker
                TotalVol = 0
                StartRow = i + 1
                OutputRow = OutputRow + 1
            Else
                
                ' find vol for same ticker
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' apply coniditional formatting for the quarterly change
        With ws.Range(ws.Cells(2, 10), ws.Cells(OutputRow - 1, 10)) ' column J
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.ColorIndex = 4 ' positive, so green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.ColorIndex = 3 ' negative, so red
        End With
        
        With ws
            ' where output goes for increase
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = TickerMaxInc
            .Cells(2, 17).Value = Format(MaxInc, "0.00%")
            
            ' where output goes for decrease
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = TickerMaxDec
            .Cells(3, 17).Value = Format(MaxDec, "0.00%")
            
            ' where output goes for volume
            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = TickerMaxVol
            .Cells(4, 17).Value = MaxVol
        End With
    Next ws
    

End Sub
