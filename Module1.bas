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
            
            
        Next i
End Sub
