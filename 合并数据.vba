Sub 合并数据()
    Dim rngSelection As Range
    Dim rngCol As Range
    Dim rngCell As Range
    Dim strCombine As String
    Dim firstCell As Range
    
    ' 如果选中对象不是单元格区域，则退出
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选择要合并的单元格区域", vbExclamation
        Exit Sub
    End If
    
    ' 将选区赋值给 rngSelection
    Set rngSelection = Selection
    
    ' 遍历选区内的每一列
    For Each rngCol In rngSelection.Columns
        ' 取该列选区的第一行单元格
        Set firstCell = rngCol.Cells(1)
        
        ' 每次遍历前，先清空拼接字符串
        strCombine = ""
        
        ' 遍历该列被选中的每一个单元格
        For Each rngCell In rngCol.Cells
            ' 跳过空白单元格
            If Trim(rngCell.Value) <> "" Then
                If strCombine = "" Then
                    ' 拼接字符串初始赋值
                    strCombine = rngCell.Value
                Else
                    ' 后续单元格的值以“、”分隔拼接
                    strCombine = strCombine & "、" & rngCell.Value
                End If
            End If
        Next rngCell
        
        ' 将拼接后的值写回本列选区最顶端的单元格
        firstCell.Value = strCombine
    Next rngCol
End Sub

