Attribute VB_Name = "Module1"

Sub stockmarket()

    For Each ws In Worksheets

        Dim Ticker As String
        Dim Totalstock As Double
        Dim openyear As Variant
        Dim closeyear As Variant
        Dim Changeday As Variant
        Dim PercentChange As Variant
        Dim maxvalue As Variant
        Dim minvalue As Variant
        Dim maxtotalvalue As Variant
        Dim rownumbermax As Variant
        Dim rownumbermin As Variant
        Dim rownumbermaxtotal As Variant
        
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        openyear = ws.Cells(2, 3).Value
        
        Totalstock = 0
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"



      For i = 2 To LastRow

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value

      closeyear = ws.Cells(i, 5).Value
      
      Yearlychange = closeyear - openyear
      
      If openyear = 0 Then
       openyear = 0.0000001
      End If
      
      PercentChange = (Yearlychange / openyear)
      
      Totalstock = Totalstock + ws.Cells(i, 7).Value

      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ws.Range("J" & Summary_Table_Row).Value = Yearlychange
      
      ws.Range("K" & Summary_Table_Row).Value = PercentChange
      
      ws.Range("K" & Summary_Table_Row).Style = "percent"
      
      ws.Range("L" & Summary_Table_Row).Value = Totalstock

      Summary_Table_Row = Summary_Table_Row + 1
      
      openyear = ws.Cells(i + 1, 3).Value
      Totalstock = 0

    Else

      Totalstock = Totalstock + ws.Cells(i, 7).Value

    End If
    
      If ws.Cells(i, 10).Value < 0 Then
       ws.Cells(i, 10).Interior.ColorIndex = 3
      Else
       ws.Cells(i, 10).Interior.ColorIndex = 4
      End If

    Next i
        
        ws.Cells(3, 15).Value = "Greatest % increase"
        ws.Cells(4, 15).Value = "Greatest % decrease"
        ws.Cells(5, 15).Value = "Greatest total volume"
        ws.Cells(2, 16).Value = "Ticker"
        ws.Cells(2, 17).Value = "Value"
        maxvalue = ws.Application.WorksheetFunction.Max(ws.Columns("K"))
        ws.Cells(3, 17).Value = maxvalue
        ws.Cells(3, 17).Style = "percent"
        minvalue = ws.Application.WorksheetFunction.Min(ws.Columns("K"))
        ws.Cells(4, 17).Value = minvalue
        ws.Cells(4, 17).Style = "percent"
        maxtotalvalue = ws.Application.WorksheetFunction.Max(ws.Columns("L"))
        ws.Cells(5, 17).Value = maxtotalvalue
        
        rownumbermax = ws.Application.WorksheetFunction.Match(ws.Cells(3, 17), ws.Range("K:K"), 0)
        ws.Cells(3, 16).Value = ws.Cells(rownumbermax, 9).Value
        
        rownumbermin = ws.Application.WorksheetFunction.Match(ws.Cells(4, 17), ws.Range("K:K"), 0)
        ws.Cells(4, 16).Value = ws.Cells(rownumbermin, 9).Value
        
        rownumbermaxtotal = ws.Application.WorksheetFunction.Match(ws.Cells(5, 17), ws.Range("L:L"), 0)
        ws.Cells(5, 16).Value = ws.Cells(rownumbermaxtotal, 9).Value

        


    Next ws


End Sub



