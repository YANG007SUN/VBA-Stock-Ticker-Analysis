'------------------------------------------------------------------------------------------------------------
'                                                                                                           '
'      This program analyzes the stock price across different year, and it gives you yearly change,         '
'      percentage change, and total stock volume for each stock. At the end. Also, it will return ticker    '
'      of the stock with greatest percentage increase, grestest percentage decrease, and greatest volume.   '
'                                                                                                           '
'------------------------------------------------------------------------------------------------------------




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
    
    
    'count number of worksheets in the workbook
    worksheet_count = Application.Worksheets.Count
    
    'loop through each worksheets
    For i = 1 To worksheet_count

            Set ws = ThisWorkbook.Worksheets(i)
            row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
 

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
            
                    'loop through new table to print out needed information
                    For j = 2 To new_row_count
                            
                            'filter to certain ticker and sort date column
                            ws.Range("A1:G" & row_count).AutoFilter field:=1, Criteria1:=ws.Cells(j, 9)
                            ws.Range("A1:G" & row_count).Sort Key1:=ws.Range("B1"), Order1:=xlAscending, Header:=xlYes
                            
                            'looking for intersected row number and last row number
                            n_row = Intersect(ws.Range("A1").CurrentRegion, ws.Range("A2:G" & row_count)).SpecialCells(xlCellTypeVisible).Row
                            n_last_row = ws.Range("A1", ws.Range("A1").End(xlDown)).End(xlDown).Row
                            
                            
                            'color yearly price based on different condition
                            If ws.Cells(n_last_row, 6) - ws.Cells(n_row, 3) < 0 Then
                            
                                ws.Cells(j, 10) = ws.Cells(n_last_row, 6) - ws.Cells(n_row, 3) 'yearly price change
                                ws.Cells(j, 10).Interior.ColorIndex = 3
                            
                            Else
                                
                                ws.Cells(j, 10) = ws.Cells(n_last_row, 6) - ws.Cells(n_row, 3) 'yearly price change
                                ws.Cells(j, 10).Interior.ColorIndex = 4
                            
                            End If
                            
                            ws.Cells(j, 11) = Format((ws.Cells(n_last_row, 6) - ws.Cells(n_row, 3)) / ws.Cells(n_row, 3), "0%") 'percent price change
                            ws.Cells(j, 12) = WorksheetFunction.Sum(ws.Range("G:G").SpecialCells(xlCellTypeVisible)) 'total stock volume
                            
                            'checking for greatest percent increase , decrease and greatest volumne
                            If greatest_percent_increase < ws.Cells(j, 11) Then
                                greatest_percent_increase = ws.Cells(j, 11)
                                ticker_greatest_percent_increase = ws.Cells(j, 9)
                            End If
                                
                        
                            If greatest_percent_decrease > ws.Cells(j, 11) Then
                                greatest_percent_decrease = ws.Cells(j, 11)
                                ticker_greatest_percent_decrease = ws.Cells(j, 9)
                            End If
                            
                            If greatest_volume < ws.Cells(j, 12) Then
                                greatest_volume = ws.Cells(j, 12)
                                ticker_greatest_volume = ws.Cells(j, 9)
                            End If
                            
                            
                    Next j
                    
                    
                    
                    'ticker name
                    ws.Cells(2, 16) = ticker_greatest_percent_increase
                    ws.Cells(3, 16) = ticker_greatest_percent_decrease
                    ws.Cells(4, 16) = ticker_greatest_volume
                    
                    'value
                    ws.Cells(2, 17) = greatest_percent_increase
                    ws.Cells(3, 17) = Format(greatest_percent_decrease, "0%")
                    ws.Cells(4, 17) = greatest_volume
                    
                    ws.AutoFilter.ShowAllData
    
    Next i
    



End Sub


