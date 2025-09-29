Sub 隔行插入空白行()
    Dim lastRow As Long
    Dim i As Long
    
    ' 将计算模式设置为手动
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False ' 关闭屏幕刷新，加快执行速度
    
    ' 获取工作表的最后一行
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 从最后一行开始向上遍历，每隔一行插入一个空行，确保公式不会被破坏
    For i = lastRow To 2 Step -1
        ' 检查当前行是否包含公式
        If Not Cells(i, 1).HasFormula Then
            Rows(i).Insert Shift:=xlDown
        End If
    Next i
    
    ' 恢复计算模式为自动
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True ' 恢复屏幕刷新
End Sub
