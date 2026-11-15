Attribute VB_Name = "ScholarshipsAmount"
Sub createSheet()
    ThisWorkbook.Worksheets.Add
    ActiveSheet.name = "Chart"
End Sub
'create a table and chart to summarise the amount of scholarships awarded over years
Sub LineChart()
    Dim ws1 As Worksheet: Dim ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Chart")
    
    Dim cols As Variant
    cols = Array(1, 2, 4, 6, 8, 10)
    
    Dim c As Integer, r As Integer
    
    For c = 1 To 6
        For r = 1 To 15
            If ws1.Cells(r, cols(c - 1)) = "-" Then
                ws2.Cells(r, c).ClearContents
            Else
            ws2.Cells(r, c) = ws1.Cells(r, cols(c - 1))
            End If
            Next r
    Next c
    
    Dim dataRange As Range
    Set dataRange = ws2.Range("A1:F15")
    
    Dim chartObj As ChartObject
    Set chartObj = ws2.ChartObjects.Add(Left:=0, Top:=ws2.Range("A17").Top, Width:=540, Height:=285)
    
    Dim colName As Variant
    colName = Array("Year", _
        "Outstanding Performance", _
        "Best Progress", _
        "Talent Development", _
        "Reaching Out", _
        "Endeavour")
    
    For i = 1 To 6
        ws2.Cells(1, i).Value = colName(i - 1)
        ws2.Cells(1, i).Font.Bold = True
        Next i

    For i = 2 To 15
        Dim cel As Variant
        cel = ws2.Cells(i, 1).Value
        ws2.Cells(i, 1).NumberFormat = "@"
        ws2.Cells(i, 1).Value = Right(cel, Len(cel) - 2)
        Next i
    
   
    For i = 2 To 14
        If i Mod 2 = 0 Then
            ws2.Range(ws2.Cells(i + 1, 1), ws2.Cells(i + 1, 6)).Interior.Color = RGB(230, 230, 230)
        End If
    Next i
        
    ws2.Columns("A").ColumnWidth = 10
    ws2.Columns("B:F").ColumnWidth = 15
    
    With ws2.Range("A1:F15")
        .Borders.LineStyle = xlNone
        .BorderAround LineStyle:=xlContinuous
    End With
    With ws2.Range("A1:A15")
        .BorderAround LineStyle:=xlContinuous
    End With
    With ws2.Range("A1:F1")
        .BorderAround LineStyle:=xlContinuous
        .Font.Color = vbWhite
        .Interior.Color = &H212121
    End With
    With ws2.Range("B1:F15")
        .HorizontalAlignment = xlCenter
    End With
    
    
    With chartObj.Chart
        .ChartType = xlLineMarkers
        .SetSourceData Source:=dataRange
        .PlotBy = xlColumns
        
        .DisplayBlanksAs = xlNotPlotted
        .Axes(xlCategory).CategoryNames = ws2.Range("A2:A15")
        
        .HasTitle = True
        .ChartTitle.Text = "Amount of Scholarships Awarded by Category ($ Million)"
        .ChartTitle.Font.Size = 14
    End With
End Sub

