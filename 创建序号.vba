Sub 创建序号()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim targetColumn As Range
    Dim maxNumber As Double
    Dim lastRow As Long
    
    ' 获取当前选定的单元格
    Set selectedCell = Selection
    
    ' 获取工作表对象
    Set ws = selectedCell.Worksheet
    
    ' 获取选定单元格所在列的最后一行
    lastRow = ws.Cells(ws.Rows.Count, selectedCell.Column).End(xlUp).Row
    
    ' 定义该列的范围（从第一行到最后一行）
    Set targetColumn = ws.Range(ws.Cells(1, selectedCell.Column), ws.Cells(lastRow, selectedCell.Column))
    
    ' 查找该列中的最大值
    On Error Resume Next
    maxNumber = Application.WorksheetFunction.Max(targetColumn)
    On Error GoTo 0
    
    ' 在选定的单元格中生成最大值+1
    If IsNumeric(maxNumber) Then
        selectedCell.Value = maxNumber + 1
    Else
        selectedCell.Value = 1 ' 如果列中没有数字，则生成1
    End If
End Sub

