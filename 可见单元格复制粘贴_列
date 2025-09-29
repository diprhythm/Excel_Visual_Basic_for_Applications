Sub 可见单元格复制粘贴()
    Dim sourceRange As Range, targetStart As Range
    Dim sourceVisibleRows As Range
    Dim wsTarget As Worksheet
    Dim colCount As Long
    Dim srcRow As Range
    Dim srcList As Collection
    Dim i As Long, pastedCount As Long
    Dim nextPasteRow As Long

    ' 选择复制区域
    On Error Resume Next
    Set sourceRange = Application.InputBox("请选择要复制的区域（可见行将被复制）", "复制区域", Type:=8)
    On Error GoTo 0
    If sourceRange Is Nothing Then MsgBox "操作取消": Exit Sub

    ' 选择粘贴起点区域
    On Error Resume Next
    Set targetStart = Application.InputBox("请选择目标区域的粘贴起点（左上角单元格）", "粘贴起点", Type:=8)
    On Error GoTo 0
    If targetStart Is Nothing Then MsgBox "操作取消": Exit Sub

    colCount = sourceRange.Columns.Count
    Set wsTarget = targetStart.Worksheet

    ' 获取源区域可见行
    On Error Resume Next
    Set sourceVisibleRows = sourceRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If sourceVisibleRows Is Nothing Then MsgBox "复制区域没有可见行": Exit Sub

    ' 收集源可见行
    Set srcList = New Collection
    For Each srcRow In sourceVisibleRows.Rows
        srcList.Add srcRow
    Next

    ' 粘贴到目标区域的可见行
    pastedCount = 0
    Dim r As Long
    r = targetStart.Row
    Do While pastedCount < srcList.Count And r <= wsTarget.Rows.Count
        If Not wsTarget.Rows(r).Hidden Then
            wsTarget.Range(wsTarget.Cells(r, targetStart.Column), _
                           wsTarget.Cells(r, targetStart.Column + colCount - 1)).Value = srcList(pastedCount + 1).Value
            pastedCount = pastedCount + 1
        End If
        r = r + 1
    Loop

    ' 如果源数据还有剩余，粘贴到区域外（未筛选区域）
    If pastedCount < srcList.Count Then
        nextPasteRow = wsTarget.Cells(wsTarget.Rows.Count, targetStart.Column).End(xlUp).Row + 1
        For i = pastedCount + 1 To srcList.Count
            wsTarget.Range(wsTarget.Cells(nextPasteRow, targetStart.Column), _
                           wsTarget.Cells(nextPasteRow, targetStart.Column + colCount - 1)).Value = srcList(i).Value
            nextPasteRow = nextPasteRow + 1
        Next i
    End If

    MsgBox "粘贴完成", vbInformation
End Sub

