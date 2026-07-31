' Calcul du total directement en VBA (indépendant du recalcul Excel)
For c = firstMonthCol To nbCols
    Dim totalVal As Double
    totalVal = Application.WorksheetFunction.Sum( _
        shDest.Range(shDest.Cells(firstDataRow, c), shDest.Cells(lastDataRow, c)))
    
    If totalVal >= 7.2 And totalVal <= 7.5 Then
        shDest.Cells(summaryRow, c).Font.Color = RGB(0, 160, 80)
    Else
        shDest.Cells(summaryRow, c).Font.Color = RGB(255, 0, 0)
    End If
Next c
