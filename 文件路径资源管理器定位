'自动运行
Sub Auto_Open()
    AddRightClickMenu
End Sub

'—— 添加右键菜单项 ——
Sub AddRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("定位到该文件").Delete
    On Error GoTo 0

    Dim btn As CommandBarButton
    Set btn = Application.CommandBars("Cell") _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    With btn
        .Caption = "定位到该文件"
        .OnAction = "LocateFile"
        .BeginGroup = True
    End With
End Sub

'—— 右键点击时定位文件 ——
Sub LocateFile()
    Dim filePath As String
    filePath = Trim(ActiveCell.Value)
    If Len(Dir(filePath)) > 0 Then
        Shell "explorer.exe /select,""" & filePath & """", vbNormalFocus
    Else
        MsgBox "文件不存在或路径不正确：" & vbCrLf & filePath, vbExclamation
    End If
End Sub


