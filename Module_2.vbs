Sub Module_2()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        
        'Set Required Variables
        Dim ticker  As String
        Dim stock   As Double
        Dim opening_price As Double
        Dim closing_price As Double
        Dim lastrow As String
        Dim summary_table_row As Integer
        Dim price_change  As Double
        Dim percentage As Double

        'Set Bonus Variables
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_volume As Double
        
        'Set Header Row for Summary Table
        Range("K1").Value = "Ticker"
        Range("L1").Value = "Yearly Change"
        Range("M1").Value = "Percent Change"
        Range("N1").Value = "Stock Volume"
        
        'Set Starting Row Value
        summary_table_row = 2
        'Set row count for range
        lastrowA = Range("A" & Rows.Count).End(xlUp).Row
        'Set stock total
        stock = 0
        
        'Iterate through each row in existing data
        For i = 2 To lastrowA

            'Sum Stock Volume for Selected Ticker
            stock = stock + Cells(i, 7)
            
            'Check If Ticker Name Different to Previous Row
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                'Set Opening Price
                opening_price = Cells(i, 3)

                'Check If Ticker Name Different to Next Row
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Print Ticker Name in New Range
                ticker = Cells(i, 1).Value
                Range("K" & summary_table_row).Value = ticker
                
                'Set Closing Price
                closing_price = Cells(i, 6)

                'Post Yearly Price Change in New Range
                price_change = closing_price - opening_price
                Range("L" & summary_table_row).Value = price_change

                'Post Percentage Changed
                percentage = (price_change / opening_price)
                Range("M" & summary_table_row).Value = Format(percentage, "Percent")
                
                'Post Stock Volume and Clear Variable
                Range("N" & summary_table_row).Value = stock
                
                'Reset Variables
                stock = 0
                price_change = 0
                opening_price = 0
                closing_price = 0
                percentage = 0
                
                'Move to Next Row
                summary_table_row = summary_table_row + 1
            End If
        Next i

        'Apply Conditional Formatting
        lastrowK = Range("K" & Rows.Count).End(xlUp).Row
        'Iterate for each colour
        For j = 2 To lastrowK
            If Cells(j, 12).Value < 0 Then
                Cells(j, 12).Interior.ColorIndex = 3
            Else
                Cells(j, 12).Interior.ColorIndex = 4
            End If
        Next j

        '----BONUS----
        'SetHeaders for Bonus
        Range("Q2").Value = "Greatest % increase"
        Range("Q3").Value = "Greatest % decrease"
        Range("Q4").Value = "Greatest total volume"
        Range("R1").Value = "Ticker"
        Range("S1").Value = "Value"

        'Set % Ranges
        lastrowM = Range("M" & Rows.Count).End(xlUp).Row

        'Set % Variables according to Max Increase/Decrease
        max_increase = WorksheetFunction.Max(Range("M1:M" & lastrowM))
        max_decrease = WorksheetFunction.Min(Range("M1:M" & lastrowM))
        Range("S2").Value = Format(max_increase, "Percent")
        Range("S3").Value = Format(max_decrease, "Percent")

        'Set Volume Range
        lastrowN = Range("N" & Rows.Count).End(xlUp).Row

        'Set Volume Range according to Max Volume
        max_volume = WorksheetFunction.Max(Range("N1:N" & lastrowN))
        Range("S4").Value = max_volume

        'Iterate through each row in the New Range
        For k = 2 To lastrowM

            'Set "Greatest % increase" Ticker Name
            If Cells(k, 13) = max_increase Then
                Range("R2").Value = Cells(k, 11)

            'Set "Greatest % decrease" Ticker Name
            ElseIf Cells(k, 13) = max_decrease Then
                Range("R3").Value = Cells(k, 11)

            'Set "Greatest total volume" Ticker Name
            ElseIf Cells(k, 14) = max_volume Then
                Range("R4").Value = Cells(k, 11)
            End If
        Next k

        'Reset Bonus Variable
        max_increase = 0
        max_decrease = 0
        max_volume = 0

    'Move to Next Worksheet
    Next

End Sub
