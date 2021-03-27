Attribute VB_Name = "Module1"
Sub basic_solution() 'My first attempt at a solution


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
    
    'MsgBox ("Beginning Sheet: " & ws.Name & vbNewLine & "Last Used Row: " & Str(lastRow) & vbNewLine & "Last Used Column: " & Str(lastColumn))
    
    'Build Summary Headers
    Cells(1, lastColumn + 2) = "Ticker"
    Cells(1, lastColumn + 3) = "Yearly Change"
    Cells(1, lastColumn + 4) = "Percent Change"
    Cells(1, lastColumn + 5) = "Total Stock Volume"
    
    'Initialize Ticker Variables
    curr_total_volume = 0
    curr_num_tickers = 1
    curr_ticker = Cells(2, 1).Value
    curr_open = Cells(2, 3).Value
    
    
    For i = 2 To lastRow
        If (Cells(i, 1).Value = curr_ticker) Then
            curr_total_volume = curr_total_volume + Cells(i, 7).Value
            curr_close = Cells(i, 6).Value
            
        Else
            'Set Values in Summary Table
            Cells(curr_num_tickers + 1, lastColumn + 2) = curr_ticker
            Cells(curr_num_tickers + 1, lastColumn + 3) = curr_close - curr_open
            
            If curr_open = 0 Then ' Catch Divisibility errors
                Cells(curr_num_tickers + 1, lastColumn + 4) = 0
            Else
                Cells(curr_num_tickers + 1, lastColumn + 4) = (curr_close - curr_open) / curr_open
            End If
                
            Cells(curr_num_tickers + 1, lastColumn + 5) = curr_total_volume
            'Format Summary Table
            If curr_close - curr_open > 0 Then 'Conditional Formatting
                Cells(curr_num_tickers + 1, lastColumn + 3).Interior.ColorIndex = 4
            ElseIf curr_close - curr_open < 0 Then
                Cells(curr_num_tickers + 1, lastColumn + 3).Interior.ColorIndex = 3
            End If
            Cells(curr_num_tickers + 1, lastColumn + 4).NumberFormat = "0.00%" 'Percentage Formatting
            
            curr_ticker = Cells(i, 1).Value 'New Ticker
            curr_open = Cells(i, 3).Value
            curr_close = Cells(i, 6).Value
            curr_total_volume = Cells(i, 7).Value 'New Volume
            curr_num_tickers = curr_num_tickers + 1 'New number of tickers
            
        End If

        Next i 'Go through row by row
    
    
    
    
    
    Next ws_index ' End of this sheet, on to next
    
    
End Sub
