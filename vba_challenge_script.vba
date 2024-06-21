Sub Activity2Macro()
    For Each ws In Worksheets
        Dim keep_going As Boolean
        Dim i As Long

        Dim current_ticker As String
        Dim previous_ticker As String
        Dim previous_ticker_group As String
        Dim ticker_group_number As Integer


        Dim first_open_price As Double
        Dim last_close_price As Double

        Dim quarterly_change As Double
        Dim quarterly_percent_change As Double
        Dim total_stock_volume As Double

        Dim greatest_percent_increase_ticker As String
        Dim greatest_percent_increase_value As Double
        Dim greatest_percent_decrease_ticker As String
        Dim greatest_percent_decrease_value As Double
        Dim greatest_total_volume_ticker As String
        Dim greatest_total_volume_value As Double


        keep_going = True
        i = 0
        ticker_group_number = 0
        total_stock_volume = 0
        greatest_percent_increase_ticker = ""
        greatest_percent_increase_value = 0
        greatest_percent_decrease_ticker = ""
        greatest_percent_decrease_value = 0
        greatest_total_volume_ticker = ""
        greatest_total_volume_value = 0

        While keep_going = True
            current_ticker = ws.Cells(i + 2, 1).Value
            previous_ticker = ws.Cells(i + 1, 1).Value
            If current_ticker <> previous_ticker Then
                previous_ticker_group = previous_ticker
                If ticker_group_number > 0 Then
                    
                    last_close_price = ws.Cells(i + 1, 6).Value
                    ws.Cells(ticker_group_number + 1, 9).Value = previous_ticker_group
                    quarterly_change = last_close_price - first_open_price
                    ws.Cells(ticker_group_number + 1, 10).Value = quarterly_change
                    quarterly_percent_change = (last_close_price - first_open_price)/ first_open_price
                    ws.Cells(ticker_group_number + 1, 11).Value = quarterly_percent_change
                    ws.Cells(ticker_group_number + 1, 12).Value = total_stock_volume

                    If greatest_percent_increase_value < quarterly_percent_change Then
                        greatest_percent_increase_value = quarterly_percent_change
                        greatest_percent_increase_ticker = previous_ticker_group
                    End If

                    If greatest_percent_decrease_value > quarterly_percent_change Then
                        greatest_percent_decrease_value = quarterly_percent_change
                        greatest_percent_decrease_ticker = previous_ticker_group
                    End If

                    If greatest_total_volume_value < total_stock_volume Then
                        greatest_total_volume_value = total_stock_volume
                        greatest_total_volume_ticker = previous_ticker_group
                    End If
                    

                Else
                    
                    ws.Cells(ticker_group_number + 1, 9).Value = "Ticker"
                    ws.Cells(ticker_group_number + 1, 10).Value = "Quarterly Change"
                    ws.Cells(ticker_group_number + 1, 11).Value = "Percent Change"
                    ws.Cells(ticker_group_number + 1, 12).Value = "Total Stock Volume"
                End If
                first_open_price = ws.Cells(i + 2, 3).Value
                ticker_group_number = ticker_group_number + 1
                total_stock_volume = 0
            End If
            total_stock_volume = total_stock_volume + ws.Cells(i + 2, 7).Value
            If current_ticker <> "" Then

            Else
                keep_going = False
            End If

            i = i + 1

        Wend
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = greatest_percent_increase_ticker
        ws.Cells(2, 17).Value = greatest_percent_increase_value
        ws.Cells(3, 16).Value = greatest_percent_decrease_ticker
        ws.Cells(3, 17).Value = greatest_percent_decrease_value
        ws.Cells(4, 16).Value = greatest_total_volume_ticker
        ws.Cells(4, 17).Value = greatest_total_volume_value


    Next ws
End Sub

