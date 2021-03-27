Attribute VB_Name = "Module1"
Sub hard_solution() 'My first attempt at a solution


Dim num_ws As Integer 'sheets as a variable so I can easily test a subset of sheets
Dim lastRow, lastColumn As Long 'Sheet Dimensions

Dim curr_ticker As String ' Ticker variable
Dim curr_total_volume, curr_num_tickers As Long ' Ticker variable
Dim curr_open, curr_close As Double


num_ws = Application.Sheets.Count


For ws_index = 1 To num_ws 'Loop through Sheets
    'Sheet size and used cells
    Set ws = Worksheets(ws_index)
    ws.Activate
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    lastColumn = 7 'Cells(2, Columns.Count).End(xlToLeft).Column
        
    'Build Summary Headers
    Cells(1, lastColumn + 2) = "Ticker"
    Cells(1, lastColumn + 3) = "Yearly Change"
    Cells(1, lastColumn + 4) = "Percent Change"
    Cells(1, lastColumn + 5) = "Total Stock Volume"
    
    'Build Super Summary Table
    Cells(1, lastColumn + 8) = "Ticker"
    Cells(1, lastColumn + 9) = "Value"
    Cells(2, lastColumn + 7) = "Greatest % Change"
    Cells(3, lastColumn + 7) = "Greatest % Decrease"
    Cells(4, lastColumn + 7) = "Total Volume"
    
    'Initialize Ticker Variables
    curr_total_volume = 0
    curr_num_tickers = 1
    curr_ticker = Cells(2, 1).Value
    curr_open = Cells(2, 3).Value
    
    'Initialize Super Summary Values
    Cells(2, lastColumn + 9) = 0 '% Change
    Cells(3, lastColumn + 9) = 0 '% Decrease
    Cells(4, lastColumn + 9) = 0 'Total Volume
    
    'Format Super Summary Table
    Cells(2, lastColumn + 9).NumberFormat = "0.00%"
    Cells(3, lastColumn + 9).NumberFormat = "0.00%"
    Cells(4, lastColumn + 9).NumberFormat = "0.00E+00"
    
    'Loop Through all the used rows in ws
    For i = 2 To lastRow
        
        'If The Ticker is the Same
        If (Cells(i, 1).Value = curr_ticker) Then
            curr_total_volume = curr_total_volume + Cells(i, 7).Value
            curr_close = Cells(i, 6).Value
        
        'If We Have a New Ticker
        Else
            
            'Set Values in Summary Table
            Cells(curr_num_tickers + 1, lastColumn + 2) = curr_ticker 'Ticker
            Cells(curr_num_tickers + 1, lastColumn + 3) = curr_close - curr_open 'Change
            Cells(curr_num_tickers + 1, lastColumn + 5) = curr_total_volume 'Total Volume
            
            'Catch Divisibility Errors for % change AND compare against super summary
            If curr_open = 0 Then
                Cells(curr_num_tickers + 1, lastColumn + 4) = 0
            Else
                Cells(curr_num_tickers + 1, lastColumn + 4) = (curr_close - curr_open) / curr_open
                
                'Update Greatest % changes
                If ((curr_close - curr_open) / curr_open) > Cells(2, lastColumn + 9).Value Then
                    Cells(2, lastColumn + 8) = curr_ticker
                    Cells(2, lastColumn + 9) = (curr_close - curr_open) / curr_open
                ElseIf ((curr_close - curr_open) / curr_open) < Cells(3, lastColumn + 9).Value Then
                    Cells(3, lastColumn + 8) = curr_ticker
                    Cells(3, lastColumn + 9) = (curr_close - curr_open) / curr_open
                End If
            End If
            
            'Update Greatest Total Volume
            If curr_total_volume > Cells(4, lastColumn + 9).Value Then
                Cells(4, lastColumn + 8) = curr_ticker
                Cells(4, lastColumn + 9) = curr_total_volume
            End If
            
            'Format Summary Table
            If curr_close - curr_open > 0 Then 'Conditional Formatting
                Cells(curr_num_tickers + 1, lastColumn + 3).Interior.ColorIndex = 4
            ElseIf curr_close - curr_open < 0 Then
                Cells(curr_num_tickers + 1, lastColumn + 3).Interior.ColorIndex = 3
            End If
            Cells(curr_num_tickers + 1, lastColumn + 4).NumberFormat = "0.00%" 'Percentage Formatting
            
            'Set new values for the current ticker
            curr_ticker = Cells(i, 1).Value 'New Ticker
            curr_open = Cells(i, 3).Value 'New Open
            curr_close = Cells(i, 6).Value 'This will get overwritten unless there is only one date of data
            curr_total_volume = Cells(i, 7).Value 'New Volume
            curr_num_tickers = curr_num_tickers + 1 'New number of tickers
            
        End If
        
        Next i 'Go through row by row
    
    
    
    
    Columns("A:P").AutoFit 'Make Pretty
    Next ws_index ' End of this sheet, on to next
    
    MsgBox ("All Done! Thanks :)     -Andrew")
    
End Sub
