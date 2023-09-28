# Challenge2

In the first 5 rows I had help from a peer on the structure and the "&" "!" signs and the loop through all the worksheets
i.e
 For Each Hojas In WORKSHEET:
    Dim total_rows&, input_row&, output_row&, output_row_count&, count_ticker&
    Dim total_stock^
    Dim opening_price!, closing_price!, yearly_change!, percent_change!
    Dim current_ticker$
    
And also I had help from the same peer the 5th assignment because I was lost "Add funcionallity to your script to return the stock with the "Greatest % increase, etc"
i.e
 Next input_row
 For input_row = 2 To output_row_count
 
 
            If Hojas.Cells(input_row, 12).Value > Hojas.Cells(4, 17).Value Then
                Hojas.Cells(4, 17).Value = Hojas.Cells(input_row, 12).Value
                Hojas.Cells(4, 16).Value = Hojas.Cells(input_row, 9).Value
            End If
            
            If Hojas.Cells(input_row, 11).Value < Hojas.Cells(3, 17).Value Then
               Hojas.Cells(3, 17).Value = Hojas.Cells(input_row, 11).Value
               Hojas.Cells(3, 16).Value = Hojas.Cells(input_row, 9).Value
            End If
            
            If Hojas.Cells(input_row, 11).Value > Hojas.Cells(2, 17).Value Then
                Hojas.Cells(2, 17).Value = Hojas.Cells(input_row, 11).Value
                Hojas.Cells(2, 16).Value = Hojas.Cells(input_row, 9).Value

