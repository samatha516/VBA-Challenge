Attribute VB_Name = "Stock_Analyzer"
Sub Stock_Analyzer()

    For Each ws In Worksheets
        
        Dim WorksheetName As String
        
        'Counters for going through loops
        Dim RowCounter As Long              'counter for reading each row in the dataset
        Dim BlockCounter As Long            'counter to keep track of beginning of each ticker block
        Dim TickCounter As Long             'counter for copying the ticker in column I
        
        'Last row number for <ticker> column (A) and distinct Ticker column (I)
        Dim LastRow As Long
        
        'Variables required for intermediate calculations for the first part
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim YearChange As Double
        Dim PercentChange As Double
        
        'Variables required forholding the intermdiate values for the 2nd part
        Dim GreatIncrTicker As String
        Dim GreatDecrTicker As String
        Dim GreatVolumeTicker As String
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVolume As Double
           
        'Get worksheet name
        WorksheetName = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
       '*********************************** Initial Part ***************************************
        'Set Ticker Counter to first row
        TickCounter = 2
        
        'Set block start row to 2
        BlockCounter = 2
        
        'Find the last non-blank cell in column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRow)
        
        'Loop through all rows
        For RowCounter = 2 To LastRow
            
            'Check if ticker name has changed - compare the current row and the row below.
            If ws.Cells(RowCounter + 1, 1).Value <> ws.Cells(RowCounter, 1).Value Then
                
                YearOpen = ws.Cells(BlockCounter, 3).Value
                YearClose = ws.Cells(RowCounter, 6).Value
                
                YearChange = YearClose - YearOpen
                
                If YearOpen <> 0 Then
                    PercentChange = YearChange / YearOpen
                Else
                    PercentChange = YearClose
                End If
                
                'Write tickername in Column I (column #9) before moving onto the next ticker
                ws.Cells(TickCounter, 9).Value = ws.Cells(RowCounter, 1).Value
                
                'Calculate and write yearly change in column J which is column #10
                 ws.Cells(TickCounter, 10).Value = YearChange
                
                'Conditional formating that will highlight positive change in green and negative change in red.
                If YearChange < 0 Then
                    ws.Cells(TickCounter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickCounter, 10).Interior.ColorIndex = 4
                End If
                
                'Calculate and write percent change in column K which is column #11
                ws.Cells(TickCounter, 11).Value = Format(PercentChange, "Percent")
            
                'Calculate and write total volume in column L which is column #12
                ws.Cells(TickCounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(BlockCounter, 7), ws.Cells(RowCounter, 7)))
            
                'Increase TickCounter by 1 to capture the next ticker
                TickCounter = TickCounter + 1
            
                'Set new start row of the ticker block
                BlockCounter = RowCounter + 1
                
            End If
            
        Next RowCounter
        
        '*************************** Challenge Part ************************************
        
        'Find last non-blank cell in column I
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRow)
        
        'Initialize the variables with values from the top cells
        GreatIncrTicker = ws.Cells(2, 9).Value
        GreatDecrTicker = ws.Cells(2, 9).Value
        GreatVolTicker = ws.Cells(2, 9).Value
        
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        GreatVol = ws.Cells(2, 12).Value
                      
        'Loop for summary details
        For RowCounter = 2 To LastRow
                                
            'For greatest increase check if next value is larger. if yes populare that value to GreatIncr else continue the comparison
            If ws.Cells(RowCounter, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(RowCounter, 11).Value
                GreatIncrTicker = ws.Cells(RowCounter, 9).Value
             End If
            
            'For greatest decrease check if next value is smaller. if yes populare that value to GreatDecr else continue the comparison
            If ws.Cells(RowCounter, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(RowCounter, 11).Value
                GreatDecrTicker = ws.Cells(RowCounter, 9).Value
            End If
            
            'For greatest total volume check if next value is larger. if yes populare that value to GreatVol else continue the comparison
            If ws.Cells(RowCounter, 12).Value > GreatVol Then
                GreatVol = ws.Cells(RowCounter, 12).Value
                GreatVolTicker = ws.Cells(RowCounter, 9).Value
            End If
            
        Next RowCounter
            
        'Write summary results in ws.Cells
        ws.Cells(2, 16).Value = GreatIncrTicker
        ws.Cells(3, 16).Value = GreatDecrTicker
        ws.Cells(4, 16).Value = GreatVolTicker
        ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
        ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
        ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
        
        'Adjust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
                    
    Next ws
        
End Sub
