Attribute VB_Name = "Module1"
Sub SumStock()

For Each ws In Worksheets
Dim worksheetName As String

Dim inc As Integer
inc = 2
Dim totalVolume As Double
totalVolume = 0
Dim openingPrice As Double
openingPrice = ws.Cells(2, 3).Value
Dim closingPrice As Double
closingPrice = 0

worksheetName = ws.Name

' Defined values along with the stock associated with it
Dim greatestIncrease As Double
greatestIncrease = -1.79769313486231E+308

Dim greatestDecrease As Double
greatestDecrease = 1.79769313486231E+308

Dim greatestVolume As Double
greatestVolume = -1.79769313486231E+308

    For i = 2 To 753001
        
        ' adds the total volume for each stock
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
             ' the ticker symbols
            ws.Cells(inc, 9).Value = ws.Cells(i, 1)
            
            ' yearly change
            closingPrice = ws.Cells(i, 6).Value
            ws.Cells(inc, 10).Value = closingPrice - openingPrice
            
            'color change
            If ws.Cells(inc, 10).Value <= 0 Then
                ws.Cells(inc, 10).Interior.Color = vbRed
                
            Else
                ws.Cells(inc, 10).Interior.Color = vbGreen
                
            End If
            
            ' % change
            ws.Cells(inc, 11).Value = ws.Cells(inc, 10).Value / openingPrice
            ws.Cells(inc, 11).NumberFormat = "0.00%"
            
            ' To see biggest increase
            If ws.Cells(inc, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(inc, 11).Value
                ws.Cells(2, 15) = ws.Cells(inc, 9).Value
                
            End If
            
            ' To see biggest decrease
            If ws.Cells(inc, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(inc, 11).Value
                ws.Cells(3, 15) = ws.Cells(inc, 9).Value
                
            End If
            
            ' volume
            ws.Cells(inc, 12).Value = totalVolume
            openingPrice = ws.Cells(i + 1, 3).Value

            ' To see biggest volume
            If ws.Cells(inc, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(inc, 12).Value
                ws.Cells(4, 15) = Cells(inc, 9).Value
                
            End If
            
            'resets the volume for each stock
            totalVolume = 0
            
            ' increment
            inc = inc + 1
            
        End If
        
    Next i
    
    ' Greatest % increase
    ws.Cells(2, 16).Value = greatestIncrease

    
    ' Greatest % decrease
    ws.Cells(3, 16).Value = greatestDecrease

    
    ' Greatest % volume
    ws.Cells(4, 16).Value = greatestVolume

Next

End Sub
