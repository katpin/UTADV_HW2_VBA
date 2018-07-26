Attribute VB_Name = "Module1"
Sub TotalVol()
    
    'Define the variables
    Dim i As Long
    Dim TotalVol As Long
    Dim rowCount As Long
    Dim counter As Integer
    Dim ws As Worksheet
    Dim openVal As Double
    Dim closeVal As Double
    Dim perVal As Double
    Dim inMax As Double
    Dim deMax As Double
    Dim valMax As Long
    Dim rowNo As Long
    
    'Loop through all worksheets
    For Each ws In Worksheets


        'Calculate for the last row
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Set the final totals to start on Row2
        counter = 2

        'Name the multiple columns with headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"


        'Start the check from Row2 to the Last Row
        For i = 2 To rowCount

            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                openVal = ws.Cells(i, 3).Value

        'Check to see if the following row is different as the current row
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        'If so, then put the ticker name in the 9th col
                ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value

        'Calculate the opening from the closing value for the ticker
                closeVal = ws.Cells(i, 6).Value
                ws.Cells(counter, 10).Value = closeVal - openVal
        
        'If the yearly change is positive mark green, if negative, mark red
                If closeVal > openVal Then
                    ws.Cells(counter, 10).Interior.ColorIndex = 10
                    ws.Cells(counter, 10).Font.Color = RGB(226, 239, 218)
                Else
                    ws.Cells(counter, 10).Interior.ColorIndex = 9
                    ws.Cells(counter, 10).Font.Color = RGB(255, 201, 201)
                End If

        'Check for divide by 0 error
                If openVal = 0 Then
                    perVal = 0

        'Calculate percent change
                Else
                    perVal = ws.Cells(counter, 10).Value / openVal

                End If

        'Format percent change to a percentage
                ws.Cells(counter, 11) = Format(perVal, "Percent")

        'Sum up the final value to the total and output to column 12
                Total = Total + ws.Cells(i, 7).Value
                ws.Cells(counter, 12).Value = Total

        'Reset values
                openVal = 0
                closeVal = 0
                Total = 0

        'Up the counter for the unique ticker column
                counter = counter + 1

        'If not, then move on to the next row and sum up the total
            Else
                Total = Total + ws.Cells(i, 7).Value
            End If


        Next i

    'Find the maximum values for multiple columns
        inMax = WorksheetFunction.Max(ws.Range("K:K"))
        deMax = WorksheetFunction.Min(ws.Range("K:K"))
        volMax = WorksheetFunction.Max(ws.Range("L:L"))

    'Apply max values in provided space
        ws.Range("P2") = ws.Cells(WorksheetFunction.Match(inMax, ws.Range("K:K"), 0), 9)
        ws.Range("Q2") = Format(inMax, "Percent")

        ws.Range("P3") = ws.Cells(WorksheetFunction.Match(deMax, ws.Range("K:K"), 0), 9)
        ws.Range("Q3") = Format(deMax, "percent")

        ws.Range("P4") = ws.Cells(WorksheetFunction.Match(volMax, ws.Range("L:L"), 0), 9)
        ws.Range("Q4") = volMax
        


    Next ws


    
End Sub

