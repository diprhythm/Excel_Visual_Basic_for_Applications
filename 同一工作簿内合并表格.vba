Sub 合并表格()
    Dim ws As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim wsCount As Integer
    Dim i As Integer
    Dim colCount As Integer
    Dim j As Integer
    Dim rng As Range
    Dim isLongNumber As Boolean
    Dim headerRow As Long
    Dim userInput As String
    Dim wb As Workbook

    ' 询问用户输入表头所在的行号
    userInput = InputBox("请输入表头所在的行号（例如：1）", "输入表头行号", "1")
    If Not IsNumeric(userInput) Or Val(userInput) < 1 Then
        MsgBox "输入无效，请输入一个有效的正整数行号。", vbExclamation
        Exit Sub
    End If
    headerRow = CLng(userInput)

    ' 获取当前工作簿
    Set wb = ActiveWorkbook

    ' 创建新的工作表用于合并结果
    Set wsMaster = wb.Sheets.Add
    wsMaster.Name = "MergedData"
    
    rowCount = 1
    wsCount = wb.Sheets.Count

    For i = 1 To wsCount
        Set ws = wb.Sheets(i)

        ' 忽略目标工作表
        If ws.Name <> wsMaster.Name Then
            ' 获取数据最后一行和列
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            colCount = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

            ' 复制表头
            If rowCount = 1 Then
                ws.Rows(headerRow).Copy Destination:=wsMaster.Rows(rowCount)
                rowCount = rowCount + 1
            End If

            ' 复制数据部分
            If lastRow > headerRow Then
                ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, colCount)).Copy _
                    Destination:=wsMaster.Cells(rowCount, 1)
                rowCount = rowCount + lastRow - headerRow
            End If
        End If
    Next i

    ' 检查是否有长数字并格式化为文本
    For j = 1 To colCount
        isLongNumber = False
        Set rng = wsMaster.Range(wsMaster.Cells(2, j), wsMaster.Cells(rowCount - 1, j))

        For Each cell In rng
            If IsNumeric(cell.Value) And Len(cell.Value) > 10 Then
                isLongNumber = True
                Exit For
            End If
        Next cell

        If isLongNumber Then
            rng.NumberFormat = "@"
        End If
    Next j

    MsgBox "合并完成！"
End Sub
