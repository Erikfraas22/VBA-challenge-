Attribute VB_Name = "Module1"
Sub Button2_Click()

    Dim ws As Worksheet
    Dim last_row As Long
    Dim summary_table_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim quarterly_percent_change As Double
    Dim Symbol As String
    Dim volume As Double
    Dim i As Long
    Dim maxvalue As Double
    Dim minvalue As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_increase_ticker_volume_number As String
    Dim total_volume As Double
    
    summary_table_row = 2 ' Initialize summary table row
    volume = 0

    ' Headers for summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
            
            ' Loop through each row
            For i = 2 To last_row
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    open_price = ws.Cells(i, 3).Value
                End If
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = last_row Then
                    close_price = ws.Cells(i, 6).Value
                    Symbol = ws.Cells(i, 1).Value
                    volume = volume + ws.Cells(i, 7).Value
                    
                    ' Calculate changes
                    quarterly_change = close_price - open_price
                    quarterly_percent_change = quarterly_change / open_price
                    
                    ' Output to summary table
                    ws.Cells(summary_table_row, 9).Value = Symbol
                    ws.Cells(summary_table_row, 10).Value = quarterly_change
                    ws.Cells(summary_table_row, 11).Value = quarterly_percent_change
                    ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                    ws.Cells(summary_table_row, 12).Value = volume
                    
                    summary_table_row = summary_table_row + 1
                    volume = 0 ' Reset volume for the next ticker
                Else
                    volume = volume + ws.Cells(i, 7).Value
                End If
            Next i
            
            ' Reset for greatest calculations
            maxvalue = -1 ' Set to -1 for comparison
            minvalue = 1 ' Set to 1 for comparison
            total_volume = 0
            
            For r = 2 To summary_table_row - 1
                If ws.Cells(r, 11).Value > maxvalue Then
                    maxvalue = ws.Cells(r, 11).Value
                    greatest_increase_ticker = ws.Cells(r, 9).Value
                End If
                
                If ws.Cells(r, 11).Value < minvalue Then
                    minvalue = ws.Cells(r, 11).Value
                    greatest_decrease_ticker = ws.Cells(r, 9).Value
                End If
                
                If ws.Cells(r, 12).Value > total_volume Then
                    total_volume = ws.Cells(r, 12).Value
                    greatest_increase_ticker_volume_number = ws.Cells(r, 9).Value
                End If
            Next r
            
            ' Output greatest calculations
            ws.Cells(2, 17).Value = "Greatest % Increase"
            ws.Cells(2, 18).Value = greatest_increase_ticker
            ws.Cells(2, 19).Value = maxvalue
            ws.Cells(2, 19).NumberFormat = "0.00%"
            
            ws.Cells(3, 17).Value = "Greatest % Decrease"
            ws.Cells(3, 18).Value = greatest_decrease_ticker
            ws.Cells(3, 19).Value = minvalue
            ws.Cells(3, 19).NumberFormat = "0.00%"
            
            ws.Cells(4, 17).Value = "Greatest Total Volume"
            ws.Cells(4, 18).Value = greatest_increase_ticker_volume_number
            ws.Cells(4, 19).Value = total_volume
            
            ' Color formatting
            For j = 2 To summary_table_row - 1
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
            Next j
            
            summary_table_row = 2 ' Reset summary table row for next worksheet
        End If
    Next ws
End Sub
