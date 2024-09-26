Sub TotalVol()
    
    'Define the variables
    Dim i As Long
    Dim ws As Worksheet
    Dim lastrow As Long
    
    Dim openval As Double
    Dim closeval As Double
    Dim totval As Double
       
    Dim currentrow As Integer

  
    'Loop through all ws
    For Each ws In Worksheets


        'Calculate last row for each ws
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Set initial values for variables
        currentrow = 2
        totval = 0
        
        'Put headers for multiple columns to be calculated
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"


        'look through all rows to find open value, closed value, total value for each ticker
        
        For i = 2 To lastrow

        'Determine the first open value for unique ticker
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                openval = ws.Cells(i, 3).Value

        'Check if next ticker is different
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Assign ticker to corresponding cell value
                ws.Cells(currentrow, 9).Value = ws.Cells(i, 1).Value

            'Assign closing value and calculate yearly change for the ticker
                closeval = ws.Cells(i, 6).Value
                ws.Cells(currentrow, 10).Value = closeval - openval
        
            'Assign colors for Yearly Change
                If closeval - openval > 0 Then
                    ws.Cells(currentrow, 10).Interior.ColorIndex = 4
                   
                ElseIf closeval - openval < 0 Then
                    ws.Cells(currentrow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(currentrow, 10).Interior.ColorIndex = 0
                End If

            'Calculate %Change and make sure denominator(openval) <>0
                If openval = 0 Then
                    ws.Cells(currentrow, 11) = 0
                Else
                    ws.Cells(currentrow, 11) = (closeval - openval) / openval
                End If
        
            'Format percent change to a percentage
                ws.Cells(currentrow, 11).NumberFormat = "0.00%"

            'Calculate Total Value and assigned to corresponding cell in table
                totval = totval + ws.Cells(i, 7).Value
                ws.Cells(currentrow, 12).Value = totval

            'Reset values
                openval = 0
                closeval = 0
                totval = 0

            'Up the counter for the unique ticker column
                currentrow = currentrow + 1

        'If ticker not different, then move on to the next row and sum up the total volume
            Else
            
                totval = totval + ws.Cells(i, 7).Value
                
            End If


        Next i

    'Find the maximum values for multiple columns and Lookup that value to return Ticker
    
    ws.Range("Q2") = Format(WorksheetFunction.Max(ws.Range("K:K")), "Percent")
    ws.Range("P2") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0), 9)
    
    ws.Range("Q3") = Format(WorksheetFunction.Min(ws.Range("K:K")), "Percent")
    ws.Range("P3") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0), 9)
    
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("P4") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0), 9)

        


    Next ws


    
End Sub
