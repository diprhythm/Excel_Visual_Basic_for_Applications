Sub 取消合并单元格()
    Dim cell As Range
    Dim mergedRange As Range
    Dim area As Range
    Dim selectedRange As Range
    Dim cellValue As Variant

    ' 禁用屏幕更新和事件
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 获取用户选中的范围
    Set selectedRange = Selection

    ' 遍历选中区域中的所有单元格
    For Each cell In selectedRange
        If cell.MergeCells Then
            ' 获取合并区域
            Set mergedRange = cell.MergeArea
            ' 存储合并单元格的值
            cellValue = cell.Value
            ' 取消合并单元格
            mergedRange.UnMerge
            ' 填充取消合并的单元格
            For Each area In mergedRange
                area.Value = cellValue
            Next area
        End If
    Next cell

    ' 恢复屏幕更新和事件
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
