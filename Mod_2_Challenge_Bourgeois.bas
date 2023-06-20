Attribute VB_Name = "Module1"
Sub MOD_2_Challenge()
    ' ----- USE THIS SUBROUTINE TO RUN THE MODULE 2 CHALLENGE ON THE DATASET -----

    ' alias current worksheet as "crt"
    Dim crt As Worksheet
    
    ' create variables for tracking max/min values
    ' max percent increase
    Dim mpi_name As String
    Dim mpi_value As Double
    
    ' min percent increase
    Dim mpd_name As String
    Dim mpd_value As Double
    
    ' greatest total volume
    Dim gtv_name As String
    Dim gtv_value As Double
    
    Dim tab_row As Long
    Dim sub_tot As Double
    Dim open_val As Double

    ' Loop through each worksheet
    For Each crt In Worksheets
        
        ' determine last row on current sheet
        last_row = crt.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' reset variables
        ' create temp variable for subtotal
        sub_tot = 0
        
        ' table row integer set to 2
        tab_row = 2
        
        'set mpi/mpd/gtv values to 0
        mpi_value = 0
        mpd_value = 0
        gtv_value = 0
        
        ' set sheet initial opening value
        open_val = crt.Cells(2, 3).Value
        
        ' create new column headers
        crt.Cells(1, 9).Value = "Ticker"
        crt.Cells(1, 10).Value = "Yearly Change"
        crt.Cells(1, 11).Value = "Percent Change"
        crt.Cells(1, 12).Value = "Total Stock Volume"
        crt.Cells(1, 16).Value = "Ticker"
        crt.Cells(1, 17).Value = "Value"
        crt.Cells(2, 15).Value = "Greatest % Increase"
        crt.Cells(3, 15).Value = "Greatest % Decrease"
        crt.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        ' loop through sheet
        For I = 2 To last_row
           
            ' for each row compare whether the stock ticker has the same name as the next row or not
            If crt.Cells(I + 1, 1).Value <> crt.Cells(I, 1).Value Then
                ' store the closing value of the current stock
                close_val = crt.Cells(I, 6).Value
                ' copy current stock name to new data table
                crt.Cells(tab_row, 9).Value = crt.Cells(I, 1).Value
                ' evaluate and propogate yearly stock change to new data table
                crt.Cells(tab_row, 10).Value = close_val - open_val
                
                ' Color the "Yearly Change" cells
                If crt.Cells(tab_row, 10).Value > 0 Then
                    crt.Cells(tab_row, 10).Interior.ColorIndex = 4
                
                ElseIf crt.Cells(tab_row, 10).Value < 0 Then
                    crt.Cells(tab_row, 10).Interior.ColorIndex = 3
                
                End If
                
                ' calculate percent change from open to close of that year year
                crt.Cells(tab_row, 11).Value = crt.Cells(tab_row, 10).Value / open_val
                ' Determine if this stock is has a higher or lower percent change than current min/max
                ' If so store current stock info to mpi/mpd variables
                If crt.Cells(tab_row, 11).Value > mpi_value Then
                    
                    mpi_name = crt.Cells(tab_row, 9).Value
                    mpi_value = crt.Cells(tab_row, 11).Value
                    
                ElseIf crt.Cells(tab_row, 11).Value < mpd_value Then
                
                    mpd_name = crt.Cells(tab_row, 9).Value
                    mpd_value = crt.Cells(tab_row, 11).Value
                
                End If
                
                crt.Cells(tab_row, 11).NumberFormat = "0.00%; [red]-0.00%"
                crt.Cells(tab_row, 12).Value = sub_tot
                
                ' Determine if current stock's tot volume is higher than current max
                If crt.Cells(tab_row, 12).Value > gtv_value Then
                    
                    ' Update gtv variables
                    gtv_name = crt.Cells(tab_row, 9).Value
                    gtv_value = sub_tot
                    
                End If
                
                ' increment tab_row and reset sub_tot
                tab_row = tab_row + 1
                sub_tot = 0
                'set new opening value
                If I < last_row Then
                    open_val = crt.Cells(I + 1, 3).Value
                End If
                 
            Else
                sub_tot = sub_tot + Cells(I, 7).Value
                
            End If
        
        Next I
        
        ' transfer stored min/max values to data table
        crt.Cells(2, 16).Value = mpi_name
        crt.Cells(2, 17).Value = mpi_value
        crt.Cells(2, 17).NumberFormat = "0.00%"
        
        crt.Cells(3, 16).Value = mpd_name
        crt.Cells(3, 17).Value = mpd_value
        crt.Cells(3, 17).NumberFormat = "0.00%"
        
        crt.Cells(4, 16).Value = gtv_name
        crt.Cells(4, 17).Value = gtv_value
        
        ' Format column widths
        crt.Range("J:Q").Columns.AutoFit
        
    Next

End Sub
