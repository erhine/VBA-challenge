Sub Modlue2VBAChallenge()

'loop through all sheets
For Each ws In Worksheets


    'set ticker name variable
    Dim Ticker As String

    'Set an initial variable for holding the total volume per ticker
    Dim TotalVolume As LongLong
    TotalVolume = 0

    'set variables for start and end of year values
    Dim YearOpen As Double
    Dim YearClose As Double

    'set variables for year change values
    Dim YearlyChange As Double
    Dim PercentChange As Double

    'keep track of each ticker in the summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2

    'set boolean variable to help capture opening price
    Dim YearOpenCaptured As Boolean


    ' Set value of summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set value of greatest increase/decrease table titles
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"


    'Define the last row in the worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through all stocks
    For i = 2 To LastRow
    
        'capture the opening price
        If YearOpenCaptured = False Then
        
            YearOpen = ws.Cells(i, 3).Value
            
            'lock opening price until the ticker changes.
            YearOpenCaptured = True
            
        End If
        
        'check if the ticker value has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set the ticker name
            Ticker = ws.Cells(i, 1).Value
            
            'add to the total stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'set the year end closing value
            YearClose = ws.Cells(i, 6).Value
            
            'Calculate year and percent change
            YearlyChange = YearClose - YearOpen
            PercentChange = YearlyChange / YearOpen
            
            'Print the ticker name in the summary table
            ws.Range("I" & SummaryTableRow).Value = Ticker
            
            'Print the yearly change in the summary table
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
            
            'Format Yearly Change cells as red or green depending if they're negative or positive/equal to 0
            If YearlyChange >= 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
            End If
            
            'Print the percent change in the summary table and format it into a percentage
            ws.Range("K" & SummaryTableRow).Value = FormatPercent(PercentChange)
            
            'Print the total stock volume in the summary table
            ws.Range("L" & SummaryTableRow).Value = TotalVolume
            
            'Add one to the summary table row
            SummaryTableRow = SummaryTableRow + 1
            
            'reset total stock volume
            TotalVolume = 0
            
            'switch boolean back to false
            YearOpenCaptured = False
            
            
        'If the ticker is still the same...
        Else
        
            'add to the total stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        End If
    
    Next i
            
    'Now, let's look at the Summary Table created above to get the greatest values
    
    'Define the last row in the table
    LastSummaryTableRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Initialize greatest increase, greatest decrease, and greatest volume variables
    GreatestPerInc = 0
    GreatestPerDec = 0
    GreatestVol = 0
    
    'loop through the table
    For i = 2 To LastSummaryTableRow
    
        'Find greatest percent increase and grab the associated ticker
        If ws.Range("k" & i).Value > GreatestPerInc Then
            GreatestPerInc = ws.Range("k" & i).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
        End If
        
        'Find greatest percent decrease and grab the associated ticker
        If ws.Range("k" & i).Value < GreatestPerDec Then
            GreatestPerDec = ws.Range("k" & i).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
        End If
        
        'Find greatest percent volume and grab the associated ticker
        If ws.Range("L" & i).Value > GreatestVol Then
            GreatestVol = ws.Range("L" & i).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
        End If
        
    Next i
        
        'Fill out table with values found in the above if statements
        ws.Cells(2, 17).Value = FormatPercent(GreatestPerInc)
        ws.Cells(3, 17).Value = FormatPercent(GreatestPerDec)
        ws.Cells(4, 17).Value = GreatestVol
        
Next ws

End Sub

