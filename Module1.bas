Attribute VB_Name = "Module1"
Sub StockMacro()
        
    'declaring variables
    Dim tickersymbol As String
    tickersymbol = " "
    Dim tickervolume As Double
    tickervolume = 0
    Dim openprice As Double
    openprice = 0
    Dim closeprice As Double
    closeprice = 0
    Dim pricechange As Double
    pricechange = 0
    Dim percentchange As Double
    percentchange = 0
    Dim lastrow As Long
    Dim i As Long
    Dim tabletracker As Long
    tabletracker = 2
            
    'function to find the total amount of rows
    lastrow = WorksheetFunction.CountA(Columns("A:A"))
    'MsgBox lastrow (testing)
            
    'declaring headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
            
    'declaring the first starting opening price
    openprice = Cells(2, 3).Value
                
        'main code to use dataset to find our objectives
        For i = 2 To lastrow
            
            'checks ticker cell to see if they are not equal
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'giving our variables values
                tickersymbol = Cells(i, 1).Value
                closeprice = Cells(i, 6).Value
                pricechange = closeprice - openprice
                'stop division by 0
                If openprice <> 0 Then
                    percentchange = (pricechange / openprice) * 100
                End If
                tickervolume = tickervolume + Cells(i, 7).Value
                    
                'creating table for the variables we are trying to find
                Range("I" & tabletracker).Value = tickersymbol
                Range("J" & tabletracker).Value = pricechange
                Range("K" & tabletracker).Value = (CStr(percentchange) & "%")
                Range("L" & tabletracker).Value = tickervolume
                
                'determining if positive or negative yearly change by colour (green/red if positive/negative)
                If (pricechange > 0) Then
                    Range("J" & tabletracker).Interior.ColorIndex = 4
                ElseIf (pricechange < 0) Then
                    Range("J" & tabletracker).Interior.ColorIndex = 3
                'in case something is off
                Else
                    Range("J" & tabletracker).Interior.ColorIndex = 15
                End If
                    
                'moving to next ticker / resetting variables
                tabletracker = tabletracker + 1
                openprice = 0
                closeprice = 0
                pricechange = 0
                percentchange = 0
                tickervolume = 0
                'redeclaring the next starting opening price
                openprice = Cells(i + 1, 3).Value
 
            Else
                'summation of total stock volume
                tickervolume = tickervolume + Cells(i, 7).Value
            End If
          
        Next i

End Sub
