Sub main()

Const nrow As Integer = 2, ncol As Integer = 3
Dim c(nrow, ncol) As Integer, d(ncol, nrow) As Integer, e(nrow, ncol) As Integer, f(ncol, nrow) As Integer
Dim i As Integer
Dim j As Integer
    For i = 1 To nrow
        For j = 1 To ncol
        c(i, j) = 6 * Rnd() + 1
        d(j, i) = 6 * Rnd() + 1
        e(i, j) = 6 * Rnd() + 1
        f(j, i) = 6 * Rnd() + 1
    Next j
Next i
Sheets("Sheet1").Select: Range("A2").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = c(i, j)
        ActiveCell.Offset(0, 1).Select
    Next j
    ActiveCell.Offset(1, -ncol).Select
Next i

Sheets("Sheet1").Select: Range("E2").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = e(i, j)
        ActiveCell.Offset(0, 1).Select
    Next j
    ActiveCell.Offset(1, -ncol).Select
Next i

Sheets("Sheet1").Select: Range("A6").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = d(j, i)
        ActiveCell.Offset(1, 0).Select
    Next j
    ActiveCell.Offset(-ncol, 1).Select
Next i

Sheets("Sheet1").Select: Range("E6").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = f(j, i)
        ActiveCell.Offset(1, 0).Select
    Next j
    ActiveCell.Offset(-ncol, 1).Select
Next i

Dim g(nrow, ncol) As Integer, h(ncol, nrow) As Integer
    For i = 1 To nrow
        For j = 1 To ncol
        g(i, j) = c(i, j) - e(i, j)
        h(j, i) = d(j, i) + f(j, i)
        Next j
    Next i

Sheets("Sheet1").Select: Range("I2").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = g(i, j)
        ActiveCell.Offset(0, 1).Select
    Next j
    ActiveCell.Offset(1, -ncol).Select
Next i

Sheets("Sheet1").Select: Range("I6").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.value = h(j, i)
        ActiveCell.Offset(1, 0).Select
    Next j
    ActiveCell.Offset(-ncol, 1).Select
Next i

Dim a(ncol, 1) As Integer, b(ncol, 1) As Integer
For j = 1 To ncol
a(j, 1) = 6 * Rnd() + 1
Next j

Sheets("Sheet1").Select: Range("N2").Select
For j = 1 To ncol
        ActiveCell.value = a(j, 1)
        ActiveCell.Offset(1, 0).Select
Next j
            For j = 1 To ncol
                For i = 1 To nrow
                b(i, 1) = b(i, 1) + c(i, j) * a(j, 1)
                Next i
            Next j
       

Sheets("Sheet1").Select: Range("N6").Select
For i = 1 To nrow
        ActiveCell.value = b(i, 1)
        ActiveCell.Offset(1, 0).Select
  
   
Next i

 End Sub