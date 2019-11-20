Sub varable1():

    Dim total As Double

    
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    For i = 2 To rowCount

        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            
            total = total + Cells(i, 7).Value

            
            Range("I" & 2 + j).Value = Cells(i, 1).Value

        
            Range("J" & 2 + j).Value = total

            
            total = 0

        
            j = j + 1

        
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

End Sub
