Function CountColorCells(color As Range,area As Range) As Long
    Dim cell As Range
    Dim coloredCount As Long

    Application.Volatile

    colorCount = 0

    For Each cell In area
        If cell.Interior.Color = color.Interior.Color Then
            colorCount = colorCount + 1
        End If
    Next cell

    CountColorCells = colorCount
End Function
