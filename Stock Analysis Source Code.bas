'-----------------------------------------------------------------------------------------------------------'
'                                                                                                           '
'     This program analyzes the stock price across different year, and it gives you yearly change,          '
'     percentage change, and total stock volume for each stock. At the end. Also, it will return ticker     '
'     of the stock with greatest percentage increase, grestest percentage decrease, and greatest volume     '
'
'-----------------------------------------------------------------------------------------------------------'




Sub Stock_analysis():

    'define some variables
    Dim worksheet_count As Integer
    Dim ws As Worksheet
    Dim ticker_greatest_percent_increase As String
    Dim ticker_greatest_percent_decrease As String
    Dim ticker_greatest_volume As String
    Dim row_count As Long
    Dim new_row_count As Long
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_volume As Double
    Dim ticker_row_counter As Long
    Dim iteration_tracker As Long
    
    
    'count number of worksheets in the workbook
    worksheet_count = Application.Worksheets.Count
    
    'loop through each worksheets
    For i = 1 To worksheet_count

            Set ws = ThisWorkbook.Worksheets(i)
            row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row 'last row of dataset in each worksheet
 

            'find out unique value of tickers
            ws.Range("A2:A" & row_count).Copy
            ws.Range("I2").PasteSpecial
            ws.Columns(9).RemoveDuplicates Columns:=Array(1)
            
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            
            
            'row name
            ws.Cells(2, 15) = "Greatest Percent Increase"
            ws.Cells(3, 15) = "Greatest Percent Decrease"
            ws.Cells(4, 15) = "Greatest Volume"
            ws.Cells(1, 16) = "Ticker"
            ws.Cells(1, 17) = "Value"
                    
            'find out new row count after remove duplicated tickers name
            new_row_count = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            
            greatest_percent_increase = 0
            greatest_percent_decrease = 0
            greatest_volume = 0
            ticker_row_counter = 2
            stock_volume = 0
            iteration_tracker = 0
            
                    'loop through the dataset
                    For j = 2 To row_count
                        
                            'check if open price is not 0 then we proceed
                            If ws.Cells(j, 3) <> 0 Then
                                    
                                      'check if currently ticker is same as next ticker
                                        If ws.Cells(j, 1) <> ws.Cells(j + 1, 1) Then
                                            
                                                'if different, print out ticker name, volume, yearly change.
                                                ws.Cells(ticker_row_counter, 9) = ws.Cells(j, 1) 'ticker name
                                                ws.Cells(ticker_row_counter, 12) = stock_volume + ws.Cells(j, 7) 'stock volume
                                                ws.Cells(ticker_row_counter, 10) = ws.Cells(j, 6) - ws.Cells(j - iteration_tracker, 3) 'yearly change
                                                
                                                'change color of yearly change, positive green, negative red
                                                If ws.Cells(ticker_row_counter, 10) > 0 Then
                                                    ws.Cells(ticker_row_counter, 10).Interior.ColorIndex = 4
                                                Else
                                                    ws.Cells(ticker_row_counter, 10).Interior.ColorIndex = 3
                                                End If
                                                
                                                ws.Cells(ticker_row_counter, 11) = Format((ws.Cells(j, 6) - ws.Cells(j - iteration_tracker, 3)) / ws.Cells(j - iteration_tracker, 3), "0%") 'percent price change
                                                
                                                'check greatest percent increase, decrease, and greatest volume
                                                If greatest_percent_increase < ws.Cells(ticker_row_counter, 11) Then
                                                    greatest_percent_increase = ws.Cells(ticker_row_counter, 11)
                                                    ticker_greatest_percent_increase = ws.Cells(ticker_row_counter, 9)
                                                End If
                                                    
                                            
                                                If greatest_percent_decrease > ws.Cells(ticker_row_counter, 11) Then
                                                    greatest_percent_decrease = ws.Cells(ticker_row_counter, 11)
                                                    ticker_greatest_percent_decrease = ws.Cells(ticker_row_counter, 9)
                                                End If
                                                
                                                If greatest_volume < ws.Cells(ticker_row_counter, 12) Then
                                                    greatest_volume = ws.Cells(ticker_row_counter, 12)
                                                    ticker_greatest_volume = ws.Cells(ticker_row_counter, 9)
                                                End If
                                                
                                                
                                                'reset all the values after print out ticker's inforamtion
                                                ticker_row_counter = ticker_row_counter + 1
                                                iteration_tracker = 0
                                                stock_volume = 0
                                            
                                        Else 'if they are same, keep adding values
                                            
                                                stock_volume = stock_volume + ws.Cells(j, 7)
                                                iteration_tracker = iteration_tracker + 1
                                        
                                        End If 'end of checking of current ticker is same as next ticker

                            End If ' end of checking if open price is o
                            
                             
                            
                    Next j
                    
                    'ticker name
                    ws.Cells(2, 16) = ticker_greatest_percent_increase
                    ws.Cells(3, 16) = ticker_greatest_percent_decrease
                    ws.Cells(4, 16) = ticker_greatest_volume
                    
                    'value
                    ws.Cells(2, 17) = Format(greatest_percent_increase, "0%")
                    ws.Cells(3, 17) = Format(greatest_percent_decrease, "0%")
                    ws.Cells(4, 17) = greatest_volume


    Next i
    



End Sub




