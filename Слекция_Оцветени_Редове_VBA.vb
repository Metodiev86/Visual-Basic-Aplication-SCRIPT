Sub SelektiraiOtsveteniRedove()
    Dim rng As Range
    Dim cell As Range
    Dim selectedRange As Range
    Dim targetColor As Long

    ' Задаваме диапазона, в който да търсим оцветени редове
    Set rng = ActiveSheet.UsedRange

    ' Задаваме цвета, който търсим (в RGB формат)
    targetColor = RGB(146, 208, 80) ' #92D050

    ' Цикъл през всички клетки в диапазона
    For Each cell In rng
        ' Проверяваме дали цветът на фона на клетката съвпада с търсения цвят
        If cell.Interior.Color = targetColor Then
            ' Ако клетката е оцветена, избираме реда над нея и самия оцветен ред
            If selectedRange Is Nothing Then
                Set selectedRange = Union(cell.EntireRow.Offset(-1), cell.EntireRow)
            Else
                Set selectedRange = Union(selectedRange, cell.EntireRow.Offset(-1), cell.EntireRow)
            End If
        End If
    Next cell

    ' Избираме всички избрани редове
    If Not selectedRange Is Nothing Then
        selectedRange.Select
    End If

End Sub
