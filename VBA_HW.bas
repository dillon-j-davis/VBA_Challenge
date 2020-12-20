Attribute VB_Name = "Module1"
Sub Stocks()

Dim lastrow As Double
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim stockamt As Integer


Cells(1, "K").Value = "Ticker"
Cells(1, "L").Value = "Yearly Change"
Cells(1, "M").Value = "Percent Change"
Cells(1, "N").Value = "Total Volume"
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
yearly_change = 0

total_volume = 0
stockamt = 1    'variable to indicate amount of stocks. only used for formatting purposes


'loop to check the value of ticker. If the ticker is the same, it will add to volume. If the ticker is different, it will calculate statistics(referencing the statistics at year start) and move to the next ticker
    
For i = 2 To lastrow
   
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        stockamt = stockamt + 1
        ticker = Cells(i, 1).Value
        Cells(stockamt, "K").Value = ticker
        yearly_change = Cells(i, 3) - Cells(i - 260, 3)
        Cells(stockamt, "L").Value = yearly_change
        percent_change = yearly_change / Cells(i - 260, "C")
        Cells(stockamt, "M").Value = percent_change
        If percent_change > 0 Then
            Cells(stockamt, "M").Interior.ColorIndex = 4
        Else
            Cells(stockamt, "M").Interior.ColorIndex = 3
        End If
        total_volume = total_volume + Cells(i, "G")
        Cells(stockamt, "N").Value = total_volume
        total_volume = 0  'resets volume for next ticker
        
        
    Else
        total_volume = total_volume + Cells(i, "G").Value
    End If
    Next i
        
        
    
        
    
    
        
        
    






End Sub

