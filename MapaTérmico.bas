Attribute VB_Name = "MapaTérmico"
Sub GRAF_MP()

Dim X, Y, Z

Worksheets("MAPA TÉRMICO").ChartObjects(1).Activate

For Z = 2 To 9

    If Cells(Z, 41) < 0.4 Then
        'red
        ActiveChart.FullSeriesCollection(1).Points(Z - 1).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
            
            ElseIf Cells(Z, 41) > 0.7 Then
                'green
                ActiveChart.FullSeriesCollection(1).Points(Z - 1).Format.Fill.ForeColor.RGB = RGB(155, 187, 89)
        
    Else
        'yellow
        ActiveChart.FullSeriesCollection(1).Points(Z - 1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
        
        
    End If
    
Next

Worksheets("MAPA TÉRMICO").ChartObjects(2).Activate

For Y = 11 To 18

    If Cells(Y, 41) < 0.4 Then
        'red
        ActiveChart.FullSeriesCollection(1).Points(Y - 10).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
            
            ElseIf Cells(Y, 41) > 0.7 Then
                'green
                ActiveChart.FullSeriesCollection(1).Points(Y - 10).Format.Fill.ForeColor.RGB = RGB(155, 187, 89)
        
    Else
        'yellow
        ActiveChart.FullSeriesCollection(1).Points(Y - 10).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
        
        
    End If
    
Next

Worksheets("MAPA TÉRMICO").ChartObjects(3).Activate

For X = 21 To 28

    If Cells(X, 41) < 0.4 Then
        'red
        ActiveChart.FullSeriesCollection(1).Points(X - 20).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
            
            ElseIf Cells(X, 41) > 0.7 Then
                'green
                ActiveChart.FullSeriesCollection(1).Points(X - 20).Format.Fill.ForeColor.RGB = RGB(155, 187, 89)
        
    Else
        'yellow
        ActiveChart.FullSeriesCollection(1).Points(X - 20).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
        
        
    End If
    
Next

End Sub
