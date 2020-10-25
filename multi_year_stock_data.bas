Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.

  'The ticker symbol.

  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  'The total stock volume of the stock.

'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub getTicker()
    Dim last_row As Long
    Dim opening_price As Double
    Dim i As Long
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim total_volume As Double
    Dim j As Long
    Dim greatest_inc As Double
    Dim greatest_inc_tick As String
    Dim greatest_dec As Double
    Dim greatest_dec_tick As String
    Dim greatest_tot_vol As Double
    Dim greatest_tot_tick As String
    Dim ws As Worksheet
    
        ' Loop through all of the worksheets
    For Each ws In Worksheets
        
        ' To get last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        j = 2
        greatest_inc = 0
        greatest_dec = 0
        greatest_tot_vol = 0
        total_volume = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ticker = ws.Range("A2").Value
        opening_price = ws.Range("C2").Value
        
        For i = 2 To last_row
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closing_price = ws.Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                
                
                'Conditional formatting
                If yearly_change < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                
                End If
                
                ticker = ws.Cells(i, 1).Value
                total_volume = total_volume + ws.Cells(i, 7).Value
                percent_change = (yearly_change / opening_price) * 100
                
                ' print values
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = yearly_change
                ws.Cells(j, 11).Value = percent_change
                ws.Cells(j, 12).Value = total_volume
                
                j = j + 1
             
                If greatest_inc < yearly_change Then
                    greatest_inc = yearly_change
                    greatest_inc_tick = ticker
                
                    ws.Cells(2, 16).Value = greatest_inc_tick
                    ws.Cells(2, 17).Value = greatest_inc
                
                End If
                
                If greatest_dec > yearly_change Then
                    greatest_dec = yearly_change
                    greatest_dec_tick = ticker
                    
                    ws.Cells(3, 16).Value = greatest_dec_tick
                    ws.Cells(3, 17).Value = greatest_dec
                
                End If
                
                If greatest_tot_vol < total_volume Then
                    greatest_tot_vol = total_volume
                    greatest_tot_tick = ticker
                    
                    ws.Cells(4, 16).Value = greatest_tot_tick
                    ws.Cells(4, 17).Value = greatest_tot_vol
                
                End If
                'j = j + 1
                total_volume = 0
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
    
        
            End If
        
        Next i
        
        'Conditional formatting
        'If ws.Cells(i, 10).Value < 0 Then
            'ws.Cells(i, 10).Interior.ColorIndex = 3
        'Else
            'ws.Cells(i, 10).Interior.ColorIndex = 4
        
        'End If
    
    ws.Columns("A:Z").AutoFit
    
    Next ws

End Sub

