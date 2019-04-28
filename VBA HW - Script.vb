Sub TestData_Sheet1()

For Each ws In Worksheets
    Dim WorksheetName As String
    WorksheetName = ws.Name

    Dim Ticker_Name As String

    Dim Volume_Total As Double
  
    Volume_Total = 0

    Dim Summary_Table_Row As Integer
  
    Summary_Table_Row = 2
  
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ws.Range("J1").Value = "Ticker_Name"
    ws.Range("K1").Value = "Volume_Total"

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Ticker_Name = ws.Cells(i, 1).Value
        
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
        
                ws.Range("K" & Summary_Table_Row).Value = Volume_Total
        
                Summary_Table_Row = Summary_Table_Row + 1
              
                Volume_Total = 0
        
            Else
    
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
            End If

        Next i
  
Next ws
    

End Sub
