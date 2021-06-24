Attribute VB_Name = "Module1"
Sub stocksmarket():
    'Variables
    Dim ws As Worksheet
    Dim tickername As String
    Dim tickertotal As LongLong
    Dim yearlychange As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim percentchange As Double
    Dim counter As Integer
    Dim summarytablerow As Integer
    Dim max_increase As Double
    Dim min_increase As Double
    Dim max_volume As Double
    
    'Set Starting Counter Values
    summarytablerow = 2
    tickertotal = 0
    counter = 0
    
'Loop to Create Summary Table 1
For Each ws In Worksheets

        'Last Row of Data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To last_row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Ticker Name
            tickername = ws.Cells(i, 1).Value
            'Yearly Change
            openprice = ws.Cells(i - counter, 3).Value
            closeprice = ws.Cells(i, 6).Value
            yearlychange = closeprice - openprice
            'Check for division by 0
            If openprice <> 0 Then
                percentchange = (yearlychange / openprice)
            Else
                percentchange = 0
            End If
            
            'Total Stock Volume
            tickertotal = tickertotal + ws.Cells(i, 7).Value
            
            'Summary Table
            ws.Range("I" & summarytablerow).Value = tickername
            ws.Range("J" & summarytablerow).Value = yearlychange
            ws.Range("K" & summarytablerow).Value = percentchange
            ws.Range("L" & summarytablerow).Value = tickertotal
            
            'Cumulative Addition of summarytablerow
            summarytablerow = summarytablerow + 1
            
            'Reset Counter Values
            tickertotal = 0
            counter = 0
            
            'Check Yearly Change: positive value is green and negative is red color
            If yearlychange > 0 Then
            ws.Range("J" & summarytablerow - 1).Interior.ColorIndex = 4
            Else
            ws.Range("J" & summarytablerow - 1).Interior.ColorIndex = 3
            End If
            
        Else
            tickertotal = tickertotal + Cells(i, 7).Value
            counter = counter + 1
        End If
    Next i

        'Create Summary Table 1 Labels and AutoFit Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        
        'Create Summary Table 2 Labels
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Last Row for Summary Table 2
        summary_last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Find Min and Max Values for Increase, Decrease, and Volume
        max_increase = WorksheetFunction.Max(ws.Range("K2:K" & summary_last_row))
        min_increase = WorksheetFunction.Min(ws.Range("K2:K" & summary_last_row))
        max_volume = WorksheetFunction.Max(ws.Range("L2:L" & summary_last_row))
    
        'Put Min and Max Values into Summary Table 2
        ws.Range("Q2").Value = max_increase
        ws.Range("Q3").Value = min_increase
        ws.Range("Q4").Value = max_volume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Loop to find Ticker Name for Greatest Increase, Decrease, and Total Volume
    For x = 2 To summary_last_row
        If ws.Cells(x, 11).Value = max_increase Then
            ws.Range("P2").Value = ws.Cells(x, 9).Value
        End If
        If ws.Cells(x, 11).Value = min_increase Then
            ws.Range("P3").Value = ws.Cells(x, 9).Value
        End If
        If ws.Cells(x, 12).Value = max_volume Then
            ws.Range("P4").Value = ws.Cells(x, 9).Value
        End If
    Next x
    
        'AutoFit Summary Table 2
        ws.Columns("O:Q").AutoFit
    
    'Reset Summary Table Row, Tickertotal, and Counter
    summarytablerow = 2
    tickertotal = 0
    counter = 0
    
Next ws
    
    
End Sub


