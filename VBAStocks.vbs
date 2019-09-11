Sub VBAStock():

'set variables and initial values
Dim totalVolume As Double

Dim i As Long
Dim yearChange As Single

Dim j As Integer

Dim sumTable As Long
Dim wkbk As Variant
Dim rowsc As Long
Dim pc As Single
Set wkbk = ActiveWorkbook.Worksheets


For Each ws In wkbk
    j = 0
    totalVolume = 0
    yearChange = 0
    sumTable = 2
    
    'set value for the titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'find row count (eventually for each ticker)
    rowsc = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'begin for loop
    For i = 2 To rowsc
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
            'store total volume for ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'deal with zero volume
            If totalVolume = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
                    
            Else 'find first value
                If ws.Cells(sumTable, 3) = 0 Then
                    For Value = sumTable To i
                        If ws.Cells(Value, 3).Value <> 0 Then
                            sumTable = Value
                            Exit For
                        End If
                    Next Value
                End If
            
                'Find the Yearly Change Column
                yearChange = (ws.Cells(i, 6) - ws.Cells(sumTable, 3))
                pc = (yearChange / ws.Cells(sumTable, 3) * 100)
            
                'define the next ticker
                sumTable = i + 1
                
                'put results in table
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = yearChange
                ws.Range("K" & 2 + j).Value = "%" & pc
                ws.Range("L" & 2 + j).Value = totalVolume
                
                'conditional formatting via VBA
                Select Case yearChange
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            
            End If
            
            'new ticker values
            totalVolume = 0
            yearChange = 0
            j = j + 1
        
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    
    Next i
    
Next ws
        
End Sub
