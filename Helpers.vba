'When cells are merged, only one cell gives you the value.
'This function should get the value from any cell that is part of the merge
Function GetMergedValue(location As Range)
    If location.MergeCells = True Then
        GetMergedValue = location.MergeArea(1, 1)
    Else
        GetMergedValue = location
    End If
End Function
