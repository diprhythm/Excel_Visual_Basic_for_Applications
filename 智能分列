Sub 智能分列()
    Dim selCol    As Range
    Dim sep       As String
    Dim colIndex  As Long
    Dim startRow  As Long, endRow As Long
    Dim r         As Long
    Dim fullText  As String
    Dim pos       As Long
    Dim firstPart As String, secondPart As String
    Dim userChoice As String
    
    ' —— 1. 选择列 ——
    On Error Resume Next
    Set selCol = Application.InputBox("请选择要分列的单列范围：", "选择列", Type:=8)
    On Error GoTo 0
    If selCol Is Nothing Then Exit Sub
    If selCol.Columns.Count <> 1 Then
        MsgBox "?? 请仅选择一列！", vbExclamation
        Exit Sub
    End If
    
    colIndex = selCol.Column
    
    ' —— 2. 输入分隔符 ——
    sep = InputBox("请输入用于分列的字符（汉字/英文字母/数字/符号）：", "分隔符")
    If sep = "" Then
        MsgBox "?? 未输入分隔符，已取消操作。", vbExclamation
        Exit Sub
    End If
    
    ' —— 3. 识别有效数据区域 ——
    Dim firstCell As Range, lastCell As Range
    With selCol
        Set firstCell = .Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext)
        Set lastCell = .Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious)
    End With
    If firstCell Is Nothing Then
        MsgBox "?? 所选列中没有数据，取消操作。", vbInformation
        Exit Sub
    End If
    startRow = firstCell.Row
    endRow = lastCell.Row
    
    ' —— 4. 选择拆分方式 ——
    userChoice = InputBox( _
        "请选择拆分方式：" & vbCrLf & _
        "1 = 分隔符放在左侧值后面（只按首个符号拆分）" & vbCrLf & _
        "2 = 分隔符放在右侧值前面（只按首个符号拆分）" & vbCrLf & _
        "3 = 分隔符单独占一列（只按首个符号拆分）" & vbCrLf & _
        "4 = 全部按分隔符拆分（按所有符号依次拆列）", _
        "分隔符位置选择", "1")
    If userChoice <> "1" And userChoice <> "2" And userChoice <> "3" And userChoice <> "4" Then
        MsgBox "?? 输入无效，操作已取消。", vbExclamation
        Exit Sub
    End If
    
    ' —— 5. 全部拆分 ——
    If userChoice = "4" Then
        Dim parts()      As String
        Dim maxParts    As Long
        Dim countParts  As Long
        Dim i           As Long
        
        ' 5.1 找出最大拆分段数
        For r = startRow To endRow
            fullText = CStr(Cells(r, colIndex).Value)
            parts = Split(fullText, sep)
            countParts = UBound(parts) - LBound(parts) + 1
            If countParts > maxParts Then maxParts = countParts
        Next r
        
        If maxParts <= 1 Then
            MsgBox "未检测到任何分隔符，取消操作。", vbInformation
            Exit Sub
        End If
        
        ' 5.2 插入 maxParts 列 & 设置文本格式
        Columns(colIndex + 1).Resize(, maxParts).Insert Shift:=xlToRight
        Columns(colIndex + 1).Resize(, maxParts).NumberFormat = "@"
        
        ' 5.3 按所有段依次填值
        For r = startRow To endRow
            fullText = CStr(Cells(r, colIndex).Value)
            parts = Split(fullText, sep)
            For i = LBound(parts) To UBound(parts)
                Cells(r, colIndex + 1 + i).Value = parts(i)
            Next i
        Next r
        
        MsgBox "全部拆分完成！共拆出 " & maxParts & " 段。", vbInformation
        Exit Sub
    End If
    
    ' —— 6. 只按首个符号拆分 ——
    Dim outCols As Range
    If userChoice = "3" Then
        Set outCols = Range(Columns(colIndex + 1), Columns(colIndex + 3))
    Else
        Set outCols = Range(Columns(colIndex + 1), Columns(colIndex + 2))
    End If
    outCols.Insert Shift:=xlToRight
    outCols.NumberFormat = "@"
    
    ' 6.1 遍历每行，按首个符号拆分
    For r = startRow To endRow
        fullText = CStr(Cells(r, colIndex).Value)
        pos = InStr(fullText, sep)
        
        If pos > 0 Then
            Select Case userChoice
                Case "1"  ' 分隔符放左侧尾部
                    firstPart = Left(fullText, pos + Len(sep) - 1)
                    secondPart = Mid(fullText, pos + Len(sep))
                    Cells(r, colIndex + 1).Value = firstPart
                    Cells(r, colIndex + 2).Value = secondPart
                Case "2"  ' 分隔符放右侧前部
                    firstPart = Left(fullText, pos - 1)
                    secondPart = sep & Mid(fullText, pos + Len(sep))
                    Cells(r, colIndex + 1).Value = firstPart
                    Cells(r, colIndex + 2).Value = secondPart
                Case "3"  ' 分隔符占中间一列
                    firstPart = Left(fullText, pos - 1)
                    secondPart = Mid(fullText, pos + Len(sep))
                    Cells(r, colIndex + 1).Value = firstPart
                    Cells(r, colIndex + 2).Value = sep
                    Cells(r, colIndex + 3).Value = secondPart
            End Select
        Else
            ' 未找到分隔符时，左移原值
            Cells(r, colIndex + 1).Value = fullText
            If userChoice = "3" Then
                Cells(r, colIndex + 2).Value = ""
                Cells(r, colIndex + 3).Value = ""
            Else
                Cells(r, colIndex + 2).Value = ""
            End If
        End If
    Next r
    
    MsgBox "操作完成！", vbInformation
End Sub


