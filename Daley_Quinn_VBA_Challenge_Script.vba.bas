Attribute VB_Name = "Module2"
Sub Ticker()
    'Create initial variables
    Dim i As Long
    Dim ws As Worksheet
  
    ' Variables to hold the greatest increase, decrease, & volume
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim VolumeTheGreat As Double
    Dim VolumeTheGreatTicker As String
  
    'Loop through all sheets
    For Each ws In Worksheets
        ' Print Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Greatest % Increase"
  
        'Set the max and min values
        maxIncrease = -1
        maxDecrease = 1
        VolumeTheGreat = 0
  
        ' Set variables, including one to hold ticker names and quarterly change values
        Dim abbrev As String
        Dim lastrow As Long
        Dim OpenValue As Double
        Dim CloseValue As Double
        Dim VolumeCounter As Double
        
        ' Determine Last Row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Establish row location for each ticker
        Dim TickerRow As Integer
        TickerRow = 2
        
        ' Set value of VolumeCounter to 0
        VolumeCounter = 0
        
        ' Loop through all tickers
        For i = 2 To lastrow
            ' Find Opening Value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                OpenValue = ws.Cells(i, 3).Value
                
                ' Add volume to volume counter
                VolumeCounter = ws.Cells(i, 7).Value
                
            Else
                ' Create condition to add volume for all cells not first
                VolumeCounter = VolumeCounter + ws.Cells(i, 7).Value
            End If

            ' Check if ticker is the same - if not print ticker name and find close value:
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the name of the ticker
                abbrev = ws.Cells(i, 1).Value
                
                ' Set the Close Value
                CloseValue = ws.Cells(i, 6).Value
                
                ' Print the ticker name in the Summary Column I
                ws.Range("I" & TickerRow).Value = abbrev
        
                ' Subtract Quarterly Close from Quarterly Open & Print
                ws.Range("J" & TickerRow).Value = CloseValue - OpenValue
                  
                ' Calculate and print percent change
                ws.Range("K" & TickerRow).Value = ((CloseValue / OpenValue) - 1)
                ws.Range("K" & TickerRow).NumberFormat = "0.00%"
                
                ' Print total volume in Column L
                ws.Range("L" & TickerRow).Value = VolumeCounter
                
                ' Check and update the greatest increase and decrease
                If ws.Range("K" & TickerRow).Value > maxIncrease Then
                    maxIncrease = ws.Range("K" & TickerRow).Value
                    maxIncreaseTicker = abbrev
                End If
                
                If ws.Range("K" & TickerRow).Value < maxDecrease Then
                    maxDecrease = ws.Range("K" & TickerRow).Value
                    maxDecreaseTicker = abbrev
                End If
                
                ' Check and update the greatest total volume
                If ws.Range("L" & TickerRow).Value > VolumeTheGreat Then
                    VolumeTheGreat = ws.Range("L" & TickerRow).Value
                    VolumeTheGreatTicker = abbrev
                End If
                
                'Set Conditional Formatting for Changes
                If ws.Range("K" & TickerRow) < 0 Then
                    ws.Range("K" & TickerRow).Interior.ColorIndex = 3
                ElseIf ws.Range("K" & TickerRow) > 0 Then
                    ws.Range("K" & TickerRow).Interior.ColorIndex = 4
                End If
                
                ' Add one to Ticker Row
                TickerRow = TickerRow + 1
            End If
        Next i

        ' Print the results of max/min/volume for each sheet
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = VolumeTheGreatTicker
        ws.Cells(4, 16).Value = VolumeTheGreat
        
    Next ws
End Sub
