# VBA-challenge-CODE
 
 'VB_Name ="Sheet1"

Sub StockData()
 Dim WS As Worksheet
     For Each WS In ActiveWorkbook.Worksheets
     WS.Activate
     
         LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
         
         Dim OpenPrice As Double
         Dim ClosePrice As Double
         Dim YearlyChange As Double
         Dim Ticker As String
         Dim PercenChange As Double
         Dim Volume As Double
         Volume = 0
         Dim Row As Double
         Row = 2
         Dim Column As Integer
         Column = 1
         Dim i As Long
         
         Cells(1, "I").Value = "Ticker"
         Cells(1, "J").Value = "Yearly Change"
         Cells(1, "k").Value = "Percent Change"
         Cells(1, "L").Value = "Total Stock Volume"
         
         OpenPrice = Cells(2, 3).Value
         
         For i = 2 To LastRow
         
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                 
                 Ticker = Cells(i, 1).Value
                 Cells(Row, 9).Value = Ticker
                 
                 ClosePrice = Cells(i, 6).Value
                 
                 YearlyChange = ClosePrice - OpenPrice
                 Cells(Row, 10).Value = YearlyChange
                 
                 If (OpenPrice = 0 And ClosePrice = 0) Then
                     PercentChange = 0
                 ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                     PercentChange = 1
                 Else
                     PercentChange = YearlyChange / OpenPrice
                     Cells(Row, 11).Value = PercentChange
                     Cells(Row, 11).NumberFormat = "00%"
                 End If
                 
                 Volume = Volume + Cells(i, 7).Value
                 Cells(Row, 12).Value = Volume
                 
                 Row = Row + 1
                 
                 OpenPrice = Cells(i + 1, 3)
                 
                 Volume = 0
             Else
                 Volume = Volume + Cells(i, 7).Value
             End If
         Next i
         Cells(2, 15).Value = "Greatest  %  Increase"
         Cells(3, 15).Value = "Greatest % Decrease"
         Cells(4, 15).Value = "Greatest Totale Volume"
         Cells(1, 16).Value = "Ticker"
         Cells(1, 17).Value = "Value"
         
         For Z = 2 To YCLastRow
             If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                 Cells(2, 16).Value = Cells(Z, 9).Value
                 Cells(2, 17).Value = Cells(Z, 11).Value
                 Cells(2, 17).NumberFormat = "0.00%"
             ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                 Cells(3, 16).Value = Cells(Z, 9).Value
                 Cells(3, 17).Value = Cells(Z, 11).Value
                 Cells(3, 17).NumberFormat = "0.00%"
             ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                 Cells(4, 16).Value = Cells(Z, 9).Value
                 Cells(4, 17).Value = Cells(Z, 11).Value
                 
             End If
             
     Next Z
             
                 
         
         
         
         
         Next WS
         
 End Sub
 
 
 
 'VB_Name = "Sheet2"
 
 Sub StockData()
 Dim WS As Worksheet
     For Each WS In ActiveWorkbook.Worksheets
     WS.Activate
     
         LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
         
         Dim OpenPrice As Double
         Dim ClosePrice As Double
         Dim YearlyChange As Double
         Dim Ticker As String
         Dim PercenChange As Double
         Dim Volume As Double
         Volume = 0
         Dim Row As Double
         Row = 2
         Dim Column As Integer
         Column = 1
         Dim i As Long
         
         Cells(1, "I").Value = "Ticker"
         Cells(1, "J").Value = "Yearly Change"
         Cells(1, "k").Value = "Percent Change"
         Cells(1, "L").Value = "Total Stock Volume"
         
         OpenPrice = Cells(2, 3).Value
         
         For i = 2 To LastRow
         
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                 
                 Ticker = Cells(i, 1).Value
                 Cells(Row, 9).Value = Ticker
                 
                 ClosePrice = Cells(i, 6).Value
                 
                 YearlyChange = ClosePrice - OpenPrice
                 Cells(Row, 10).Value = YearlyChange
                 
                 If (OpenPrice = 0 And ClosePrice = 0) Then
                     PercentChange = 0
                 ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                     PercentChange = 1
                 Else
                     PercentChange = YearlyChange / OpenPrice
                     Cells(Row, 11).Value = PercentChange
                     Cells(Row, 11).NumberFormat = "00%"
                 End If
                 
                 Volume = Volume + Cells(i, 7).Value
                 Cells(Row, 12).Value = Volume
                 
                 Row = Row + 1
                 
                 OpenPrice = Cells(i + 1, 3)
                 
                 Volume = 0
             Else
                 Volume = Volume + Cells(i, 7).Value
             End If
         Next i
         Cells(2, 15).Value = "Greatest  %  Increase"
         Cells(3, 15).Value = "Greatest % Decrease"
         Cells(4, 15).Value = "Greatest Totale Volume"
         Cells(1, 16).Value = "Ticker"
         Cells(1, 17).Value = "Value"
         
         For Z = 2 To YCLastRow
             If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                 Cells(2, 16).Value = Cells(Z, 9).Value
                 Cells(2, 17).Value = Cells(Z, 11).Value
                 Cells(2, 17).NumberFormat = "0.00%"
             ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                 Cells(3, 16).Value = Cells(Z, 9).Value
                 Cells(3, 17).Value = Cells(Z, 11).Value
                 Cells(3, 17).NumberFormat = "0.00%"
             ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                 Cells(4, 16).Value = Cells(Z, 9).Value
                 Cells(4, 17).Value = Cells(Z, 11).Value
                 
             End If
             
     Next Z
             
                 
         
         
         
         
         Next WS
         
 End Sub



'VB_Name = "Sheet3"

Sub StockData()
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Ticker As String
        Dim PercenChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "k").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        OpenPrice = Cells(2, 3).Value
        
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                Ticker = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker
                
                ClosePrice = Cells(i, 6).Value
                
                YearlyChange = ClosePrice - OpenPrice
                Cells(Row, 10).Value = YearlyChange
                
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / OpenPrice
                    Cells(Row, 11).Value = PercentChange
                    Cells(Row, 11).NumberFormat = "00%"
                End If
                
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                
                Row = Row + 1
                
                OpenPrice = Cells(i + 1, 3)
                
                Volume = 0
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        Cells(2, 15).Value = "Greatest  %  Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Totale Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        For Z = 2 To YCLastRow
            If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Z, 9).Value
                Cells(2, 17).Value = Cells(Z, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Z, 9).Value
                Cells(3, 17).Value = Cells(Z, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Z, 9).Value
                Cells(4, 17).Value = Cells(Z, 11).Value
                
            End If
            
    Next Z
            
                
        
        
        
        
        
       Next WS
End Sub

