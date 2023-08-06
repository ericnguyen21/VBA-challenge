Attribute VB_Name = "Module1"
Sub stock_check()
    
    
    For Each ws In Worksheets
    
    Dim row_number As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Stock_Volume"
    
    Dim ticker_code As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim stock_volume As Double
    Dim openvalue As Double
    Dim closevalue As Double
    
    stock_volume = 0
    
    Dim summary_row As Integer
    summary_row = 2
    
    Dim i As Long
    
        openvalue = ws.Cells(2, 3).Value
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                closevalue = ws.Cells(i, 6).Value
                
                ticker_code = ws.Cells(i, 1).Value
            
                yearly_change = closevalue - openvalue
            
                percentage_change = (yearly_change / openvalue) * 100
            
                stock_volume = stock_volume + ws.Cells(i, 7).Value
            
                ws.Range("I" & summary_row).Value = ticker_code
                ws.Range("J" & summary_row).Value = yearly_change
                ws.Range("K" & summary_row).Value = percentage_change
                    
                    If percentage_change < 0 Then
                        ws.Range("K" & summary_row).Interior.ColorIndex = 3
                    Else
                        ws.Range("K" & summary_row).Interior.ColorIndex = 4
                    End If
                    
                ws.Range("L" & summary_row).Value = stock_volume
            
            
                summary_row = summary_row + 1
            
                stock_volume = 0
                
                openvalue = ws.Cells(i + 1, 3).Value
                
            
            Else
            
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
            
            End If
        
        Next i
        
'Bonus
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    Dim price_increase, price_decrease As Double
    Dim total_volume As Double
    Dim bonus_lastrow As Integer
    Dim ticker_priceincrease, ticker_pricedecrease, ticker_total_volume As String
    price_increase = 0
    price_decrease = 0
    total_volume = 0
    
    bonus_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To bonus_lastrow
    
        If price_increase < ws.Cells(i, 11) Then
            price_increase = ws.Cells(i, 11)
            ticker_priceincrease = ws.Cells(i, 9)
            ws.Range("P2").Value = ticker_priceincrease
            ws.Range("Q2").Value = price_increase
        End If
        
        If price_decrease > ws.Cells(i, 11) Then
            price_decrease = ws.Cells(i, 11)
            ticker_pricedecrease = ws.Cells(i, 9)
            ws.Range("P3").Value = ticker_pricedecrease
            ws.Range("Q3").Value = price_decrease
        End If
        
        If total_volume < ws.Cells(i, 12) Then
            total_volume = ws.Cells(i, 12)
            ticker_total_volume = ws.Cells(i, 9)
            ws.Range("P4").Value = ticker_total_volume
            ws.Range("Q4").Value = total_volume
        End If
        
    Next i
       
    Next ws
    
End Sub


