Attribute VB_Name = "Module1"
Sub Calculate_Metrics()
    'Declare variables
    Dim lngLastRow As Long              '# of last row with data for current sheet
    Dim dblEODate As Double             'Date of earliest open price for current ticker
    Dim dblEOPrice As Double            'Earliest open price for current ticker
    Dim dblLCDate As Double             'Date of latest close price for current ticker
    Dim dblLCPrice As Double            'Latest close price for current ticker
    Dim dblYearlyChange As Double       'Yearly Change for current ticker
    Dim dblPercentChange As Double      'Percent Change for current ticker
    Dim dblTotalVolume As Double        'Total Volume for current ticker
    Dim intOutputRow As Integer         'Current row of output table
    Dim intLastRowSummary As Integer    '# of last row with summary data for each ticker
    Dim dblGDSummary As Double          'Greatest decrease of summary data
    Dim strGDTicker As String           'Greatest decrease ticker
    Dim dblGISummary As Double          'Greatest increase of summary data
    Dim strGITicker As String           'Greatest increase ticker
    Dim dblGVSummary As Double          'Greatest total volume
    Dim strGVTicker As String           'Greatest total volume ticker
    
    
    'Initialize values
    dblEODate = 99999999
    dblEOPrice = 0
    dblLCDate = 0
    dblLCPrice = 99999999
    dblYearlyChange = 0
    dblPercentChange = 0
    intOutputRow = 2
    dblTotalVolume = 2
    dblGDSummary = 0
    dblGISummary = 0
    dblGVSummary = 0
    
    'Outer loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        
        'Get row # of last row with data
        lngLastRow = ws.Cells.Find(What:="*", _
                After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

        'Create Header Row - Summary Data
        ws.Range("I1").Value = "Ticker Name"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        'Create table for greatest increase/decrease data
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest total volume"
        
        'Loop through rows of data
        For lngRow = 2 To lngLastRow
            'New approach - only works if data is sorted by A column, then by date!
            
            'Get earliest opening date - compare date, store if earlier than current
            If ws.Range("B" & lngRow).Value < dblEODate Then 'Earlier than stored date
                'Store new earliest open info
                dblEODate = ws.Range("B" & lngRow).Value
                dblEOPrice = ws.Range("C" & lngRow).Value
            End If
            
            'Get latest close date - compare date, store if later than current
            If ws.Range("B" & lngRow).Value > dblLCDate Then 'Later than stored date
                'Store new latest close info
                dblLCDate = ws.Range("B" & lngRow).Value
                dblLCPrice = ws.Range("F" & lngRow).Value
            End If
            
            'Add to total volume
            dblTotalVolume = dblTotalVolume + ws.Range("G" & lngRow).Value
            
            'Check if next row matches current -> end of contiguous data
            If ws.Range("A" & lngRow).Value <> ws.Range("A" & (lngRow + 1)).Value Then
                
                'Calc yearly change
                dblYearlyChange = dblLCPrice - dblEOPrice
                
                'Calc percent change
                If dblEOPrice <> 0 Then 'Prevent division by 0
                    dblPercentChange = dblYearlyChange / dblEOPrice
                Else
                    dblPercentChange = 0
                End If
                
                'Store current values on sheet
                ws.Range("I" & intOutputRow).Value = ws.Range("A" & lngRow).Value
                ws.Range("J" & intOutputRow).Value = Format(dblYearlyChange, "#,####0.0000") 'Round to make it more comprehensible and consistent
                ws.Range("K" & intOutputRow).Value = Format(dblPercentChange, "Percent") 'Format as percent
                ws.Range("L" & intOutputRow).Value = dblTotalVolume
                
                'Check to see which formatting to apply to Yearly Change
                If dblYearlyChange > 0 Then
                    ws.Range("J" & intOutputRow).Interior.ColorIndex = 4 'Green
                ElseIf dblYearlyChange < 0 Then
                    ws.Range("J" & intOutputRow).Interior.ColorIndex = 3 'Red
                Else
                    ws.Range("J" & intOutputRow).Interior.ColorIndex = 6 'Yellow
                End If
                
                'Reset values
                dblEODate = 99999999
                dblEOPrice = 0
                dblLCDate = 0
                dblLCPrice = 99999999
                dblYearlyChange = 0
                dblPercentChange = 0
                dblTotalVolume = 0
                intOutputRow = intOutputRow + 1
                
            End If
        Next lngRow
        
        'Loop through summary data
        For intRow = 2 To intOutputRow
            'Check for greatest increase
            If ws.Range("K" & intRow).Value > dblGISummary Then
                strGITicker = ws.Range("I" & intRow).Value
                dblGISummary = ws.Range("K" & intRow).Value
            End If
            
            'Check for greatest decrease
            If dblGDSummary > ws.Range("K" & intRow).Value Then
                strGDTicker = ws.Range("I" & intRow).Value
                dblGDSummary = ws.Range("K" & intRow).Value
            End If
            
            'Check for greatest volume
            If ws.Range("L" & intRow).Value > dblGVSummary Then
                strGVTicker = ws.Range("I" & intRow).Value
                dblGVSummary = ws.Range("L" & intRow).Value
            End If
            
        Next intRow
        
        
        'Output final greatest increase, decrease, volume data
        
        ws.Range("O2").Value = strGITicker
        ws.Range("O3").Value = strGDTicker
        ws.Range("O4").Value = strGVTicker
        ws.Range("P2").Value = Format(dblGISummary, "Percent")
        ws.Range("P3").Value = Format(dblGDSummary, "Percent")
        ws.Range("P4").Value = dblGVSummary
        
        'Reset variables
        dblGISummary = 0
        dblGDSummary = 0
        dblGVSummary = 0
        strGITicker = ""
        strGDTicker = ""
        strGVTicker = ""
        intOutputRow = 2
        
    Next
    
End Sub


