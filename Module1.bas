Attribute VB_Name = "Module1"
Sub StockMacro():
    
    'declaring amount of total rows
    Dim lastrow As Long
    lastrow = WorksheetFunction.CountA(Columns("A:A"))
    'MsgBox lastrow
        
        'for loop or macro
        For i = 2 To lastrow
            'declaring variables to use
            Dim tickersymbol As String
            tickersymbol = " "
            Dim totalticker As Double
            totalticker = 0
            Dim openprice As Double
            openprice = Cells(2, 3).Value
            Dim closeprice As Double
            closeprice = 0
            Dim changeprice As Double
            changeprice = 0
            Dim changepercent As Double
            changepercent As Double
            
            'declaring headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
                'real coding begins
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'setting ticker symbol
                    tickersymbol = Cells(i, 1).Value
                    'calculating the changes
                    closeprice = Cells(i, 6).Value
                    changeprice = closeprice - openprice
                    
                        If openprice <> 0 Then
                            changepercent = (changeprice / openprice) * 100
                        Else
                            changepercent = "Divided by 0"
                        End If
                            
                    'solving for total ticker volume
                    totalticker = totalticker + Cells(i, 7).Value
                    
                        'determining if positive or negative yearly change by colour (green/red if positive/negative)
                        If (Cells(i, 10).Value > 0) Then
                            Range("J").Interior.Color = 4
                        Else
                            Range("J").Interior.Color = 3
                        End If
                        
                    
        Next i
End Sub
