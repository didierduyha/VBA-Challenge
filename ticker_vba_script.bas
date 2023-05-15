Attribute VB_Name = "Module1"
Sub RunOnAll()

    'Establishing Variables
    Dim xSh As Worksheet
    
    'Runs through all the sheets and runs the Ticker code
    For Each xSh In Worksheets
    
        xSh.Select
        Call Ticker
        
    Next
    
End Sub

Sub Ticker()

    'Establishing Variables
    Dim i, n As Integer
        n = 2
    Dim op, cl, gi, gd As Double
        op = Range("C2").Value
        gi = 0
        gd = 0
    Dim LR As Long
        Volume = Range("G2").Value
    Dim ti, td, tt As String
    
    'Setting the Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Year Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stack Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greastest % Increase"
    Range("O3").Value = "Greastest % Decrease"
    Range("O4").Value = "Greastest Total Volume"
    
    'Finds the row number of the Last Row (LR)
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop Through all the Rows
    For i = 3 To LR + 1
        
        'Add Volume to existing total if the Ticker is same as Previous Row's
        If Range("A" & i).Value = Range("A" & i - 1).Value Then
            
            Volume = Volume + Range("G" & i).Value
            
        Else
        
            'Establish the closing price for the Ticker
            cl = Range("F" & i - 1).Value
            
            'Filling out Details for the Ticker
            Range("I" & n).Value = Range("A" & i - 1).Value
            Range("J" & n).Value = cl - op
            Range("J" & n).NumberFormat = "$##0.00"
            dPercent = (cl - op) / op
            Range("K" & n).Value = FormatPercent(dPercent)
            Range("L" & n).Value = Volume
            
            'Formatting the Year Change depending (Green if Positive; Red if Negative)
            If cl - op > 0 Then
            
                'Green
                Range("J" & n).Interior.ColorIndex = 4
                Range("K" & n).Interior.ColorIndex = 4
                
            Else
                
                'Red
                Range("J" & n).Interior.ColorIndex = 3
                Range("K" & n).Interior.ColorIndex = 3
                
            End If
            
            'Establishing Greatest % Increase & Decrease
            If dPercent > gi Then
            
                gi = dPercent
                ti = Range("A" & i - 1).Value
                
            ElseIf dPercent < gd Then
            
                gd = dPercent
                td = Range("A" & i - 1).Value
                
            End If
            
            'Establishing Greatest Volume Total
            If Volume > gt Then
            
                gt = Volume
                tt = Range("A" & i - 1).Value
                
            End If
            
            'Resetting starting Volume, Open Price
            op = Range("C" & i).Value
            Volume = Range("G" & i).Value
            n = n + 1
        
        End If

    Next i
    
    'Filling out Details for "Greatests"
    Range("P2").Value = ti
    Range("P3").Value = td
    Range("P4").Value = tt
    Range("Q2").Value = FormatPercent(gi)
    Range("Q3").Value = FormatPercent(gd)
    Range("Q4").Value = gt
    
End Sub
