Attribute VB_Name = "Module1"
Sub Ticker():

For Each ws In Worksheets
    
    Dim WorksheetName As String
    WorksheetName = ws.Name

    Dim i As Double
    Dim r As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim start As Double
    Dim sum As Double
    
    r = 2
    Opening_Price = ws.Cells(2, 3).Value
    
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        sum = sum + ws.Cells(i, 7).Value
        Closing_Price = ws.Cells(i, 6).Value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(r, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(r, 10).Value = Closing_Price - Opening_Price
                If Closing_Price - Opening_Price < 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 3
                ElseIf Closing_Price - Opening_Price > 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 4
                End If
            ws.Cells(r, 11).Value = (Closing_Price - Opening_Price) / Opening_Price * 100 & "%"
            ws.Cells(r, 12).Value = sum
            r = r + 1
            Opening_Price = ws.Cells(i + 1, 3).Value
            Closing_Price = 0
            sum = 0
        End If
    Next i
    
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total As Double
    
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total = 0
    
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = Greatest_Increase * 100 & "%"
        End If
    Next i

    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = Greatest_Decrease * 100 & "%"
        End If
    Next i

    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 12).Value > Greatest_Total Then
            Greatest_Total = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = Greatest_Total
        End If
    Next i
    
Next ws
    
End Sub
