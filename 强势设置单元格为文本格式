Sub 文本()
    ' 如果选中对象不是单元格区域，则直接退出
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' 关闭屏幕更新、计算和事件，以提升处理效率
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Cleanup  ' 发生错误时，跳转到清理环节，以便恢复设置
    
    ' 设置选中区域单元格格式为文本
    Selection.NumberFormat = "@"
    
    ' 将值重新赋给自己，强制刷新格式
    Selection.Value = Selection.Value
    
Cleanup:
    ' 恢复屏幕更新、计算和事件设置
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

