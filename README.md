Sub stocks()
    ' set up to loop thru all worksheets in file
 
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        'Find last row in worksheet
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Column Values
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Quarterly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
      
        'assign variables and values
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim QuarterlyChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long

        
        'What is Open Price?
        OpenPrice = Cells(2, Column + 2).Value
        
        For i = 2 To LastRow
         ' identify ticker switch, then what is ticker, close price, and quarterly change?
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                Ticker = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker
                
                ClosePrice = Cells(i, Column + 5).Value
                
                QuarterlyChange = ClosePrice - OpenPrice
                Cells(Row, Column + 9).Value = QuarterlyChange
                
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                    
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                    
                Else
                    PercentChange = QuarterlyChange / OpenPrice
                    Cells(Row, Column + 10).Value = PercentChange
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                ' what is total volume?
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' add to summary report
                Row = Row + 1
                ' reset OP
                OpenPrice = Cells(i + 1, Column + 2)
                ' reset Vol
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, Column + 6).Value
                
            End If
            
        Next i
        
        ' Find last row of Q Change
        QCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' conditional format cells for x >0, x = 0, x <0
        For j = 2 To QCLastRow
            If (Cells(j, Column + 9).Value > 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            ElseIf Cells(j, Column + 9).Value = 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 2
            End If
        Next j
        
        ' additional summary info headers
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        ' find values for additional summary info
         For k = 2 To QCLastRow
            If Cells(k, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & QCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(k, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
                
            ElseIf Cells(k, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & QCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(k, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
                
            ElseIf Cells(k, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & QCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(k, Column + 11).Value
                
            End If
            
        Next k
        
        
    
        
    Next WS
        
End Sub
