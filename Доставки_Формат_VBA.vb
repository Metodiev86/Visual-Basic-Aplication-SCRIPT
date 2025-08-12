Sub Delivery_Format()

    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("C:C,E:E,G:G,L:L,M:M, P:P").Select
    Range("M1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A2").Select
End Sub
