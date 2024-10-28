Attribute VB_Name = "Module1"
Sub StockCounter()

    Dim Total As Double
    Dim Row As Long
    Dim RowCount As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim SummaryTableRow As Long
    Dim StockStartRow As Long
    Dim StartValue As Long
    Dim LastTicker As String
    
    'loop through all worksheets in the excel workbooko
    For Each ws In Worksheets
    
        'set headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Quarterly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'set up title row of aggregate section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Voume"
        
        
        'set initial values to variables holding ticker totals and rows
        SummaryTableRow = 0
        Total = 0
        QuarterlyChange = 0
        StockStartRow = 2
        StartValue = 2
        
        'get value of the last row in current sheet
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Find the last ticker so that we can break out of the loop
        LastTicker = ws.Cells(RowCount, 1).Value
        
        'loop until we get to teh end of the sheet
        For Row = 2 To RowCount
        
            'Check to see if the ticker changed
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
            'If there is a change in column A
            
            'First add to running total
            Total = Total + ws.Cells(Row, 7).Value
            
            'Check to see if the value of hte total stock volume is -
            If Total = 0 Then
            
                'print the results in the summary table section (columns I-L)
                Ticker = ws.Cells(Row, 1).Value                                    'Prints Ticker Value
                ws.Range("I" & 2 + SummaryTableRow).Value = ws.Cells(Row, 1).Value    'Print the Ticker Symbol in the Summary
                ws.Range("J" & 2 + SummaryTableRow).Value = 0                      'Print 0 in the Quarterly Change
                ws.Range("K" & 2 + SummaryTableRow).Value = 0                      'Print 0 in total change
                ws.Range("L" & 2 + SummaryTableRow).Value = 0                      'Print 0 in Total Stock Volume
            
            Else
                'Find the first non-zerio first open value for the stock
                If ws.Cells(StartValue, 3).Value = 0 Then
                    'if the first open is 0, search for the first non-zero stock open value by moving to teh next rows
                        For FindValue = StartValue To Row
                        
                            'check to see if the next open value does not equal 0
                            If ws.Cells(FindValue, 3).Value <> 0 Then
                                'once we have a non-zero first open value, that value becomes the row we track first open from
                                StartValue = FindValue
                                Exit For
                            End If
                        
                        Next FindValue
                End If
                
                'Calculate Quarterly Change (last close - first open)
                QuarterlyChange = ws.Cells(Row, 6).Value - ws.Cells(StartValue, 3).Value
                
                'calculate the percent change (quarterly change / first open)
                PercentChange = QuarterlyChange / ws.Cells(StartValue, 3).Value
                
                'print the results in summary table section
                ws.Range("I" & 2 + SummaryTableRow).Value = ws.Cells(Row, 1).Value    'Print the Ticker Symbol in the Summary
                ws.Range("J" & 2 + SummaryTableRow).Value = QuarterlyChange        'Print Quarterly Change
                ws.Range("K" & 2 + SummaryTableRow).Value = PercentChange          'Print perecent Total
                ws.Range("L" & 2 + SummaryTableRow).Value = Total                  'Print Total
                
                'color the quarterly change column basedon value
                                
                If QuarterlyChange > 0 Then
                    'color the cell green
                    ws.Range("J" & 2 + SummaryTableRow).Interior.ColorIndex = 4
            
                ElseIf QuarterlyChange < 0 Then
                    'color the cell red
                    ws.Range("J" & 2 + SummaryTableRow).Interior.ColorIndex = 3
                    
                Else
                    'color the cell clear/no change
                    ws.Range("J" & 2 + SummaryTableRow).Interior.ColorIndex = 0
                    
                End If
                
                'Reset/update values for the next ticker
                Total = 0           'reset total
                AverageChange = 0   'reset Average Change
                QuarterlyChange = 0 'Reset Quarterly Change
                StartValue = Row + 1    'Move start row
                SummaryTableRow = SummaryTableRow + 1   'Index summary table position
                
                
            End If
            
            Else 'If we are in the same ticker, keep adding to Total
            Total = Total + ws.Cells(Row, 7).Value 'Get the value from the sevent column of current Row and add to total
            
            End If
            
                 
        Next Row
        
        
            'clean up if needed to avoid having extra data be placed in the summary section
            'find the last row of data
            
            'update summary table row
            SummaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
            
            'find the last data in the extra rows from columns J:L
            Dim LastExtraRow As Long
            LastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
            'loop that clears the extra dta
            For e = SummaryTableRow To LastExtraRow
            'for loop that goes through columns I-L (9-12)
                For Column = 9 To 12
                    ws.Cells(e, Column).Value = ""
                    ws.Cells(e, Column).Interior.ColorIndex = 0
               Next Column
            Next e
            
        'print the summary aggregates
        'after generating the info in the summary section find greatest % increase, %decrease and Total Volume
          
        SummaryTotalRows = ws.Cells(Rows.Count, "K").End(xlUp).Row
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & SummaryTotalRows))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & SummaryTotalRows))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & SummaryTotalRows))
          
        'Use Match to find tickers associated with greatest increase, decrease and total
        Dim GreatestIncreaseRow As Double
        Dim GreatestDecreaseRow  As Double
        Dim GreatestTotVolRow As Double
        
        GreatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & SummaryTableRow)), ws.Range("K2:K" & SummaryTableRow + 2), 0)
        GreatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & SummaryTableRow)), ws.Range("K2:K" & SummaryTableRow + 2), 0)
        GreatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & SummaryTableRow)), ws.Range("L2:L" & SummaryTableRow + 2), 0)
        
        'Display the ticker
        ws.Range("P2").Value = ws.Cells(GreatestIncreaseRow + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(GreatestDecreaseRow + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(GreatestDecreaseRow + 1, 9).Value
        
        'format the summary table columns
        For s = 0 To SummaryTableRow
            ws.Range("J" & 2 + s).NumberFormat = "0.00"
            ws.Range("K" & 2 + s).NumberFormat = "0.00%"
            ws.Range("L" & 2 + s).NumberFormat = "#,###"
        Next s
        
        
        'format the summary aggregates
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###"
        
        
        'Autofit the info across all columns
        ws.Columns("A:Q").AutoFit
    
    Next ws

End Sub
