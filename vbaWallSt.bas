Sub VBAWallSt()
    
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    'Declare variables
    Dim Ticker As String
    Dim LR_A As Long
    Dim LR_K As Long
    Dim SummaryTableRow As Long
    Dim OP As Double
    Dim CP As Double
    Dim Yearly_Change As Double
    Dim Previous As Long
    Dim Percent_Change As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Double
    
    'Label Columns and Tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Set Counters for Reset
    TTVolume = 0
    SummaryTable = 2
    Previous = 2
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
         
    'Find Last Row Value of A
    LR_A = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through rows
    For i = 2 To LR_A

        'Add the values into Total Ticker Vol
        TTVolume = TTVolume + ws.Cells(i, 7).Value
    
        'Check for different name values
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set Ticker Name for the first column
            TicName = ws.Cells(i, 1).Value
                
            'Summary Table
            ws.Range("I" & SummaryTable).Value = TicName
                
            ws.Range("L" & SummaryTable).Value = TTVolume
               
            'Reset Total Ticker Volume
            TTVolume = 0

            'Yearly Open
            OP = ws.Range("C" & Previous)
                
            'Yearly Close
            CP = ws.Range("F" & i)
                
            'Yearly Change
            Yearly_Change = CP - OP
            ws.Range("J" & SummaryTable).Value = Yearly_Change
            ws.Range("J" & SummaryTable).NumberFormat = "$0.00"

            'Percent Change
            If OP = 0 Then
                Percent_Change = 0
    
                Else
                OP = ws.Range("C" & Previous)
                Percent_Change = Yearly_Change / OP
                        
            End If
                
            'Percent Change
            ws.Range("K" & SummaryTable).Value = Percent_Change
                
            'Format to Percent
            ws.Range("K" & SummaryTable).NumberFormat = "0.00%"

            'Conditional Formatting
            If ws.Range("J" & SummaryTable).Value >= 0 Then
            ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                    
                Else
                
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                
            End If
            
            
            SummaryTable = SummaryTable + 1
              
            'Set Previous Amount
            Previous = i + 1
                
        End If
                
        'Go to next row
        Next i

        'Find Last Row K
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Loop through rows for final result table
        For i = 2 To LastRowK
            
            'Greatest % Increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'Greatest % Decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub


