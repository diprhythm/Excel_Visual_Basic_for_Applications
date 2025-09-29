' —— 高性能 + 仅作用于可见单元格 的 浅绿/无填充 切换（最终稳健版）——
' 使用说明：
'   1) 如需覆盖条件格式，把 skipCF 设为 False（默认 True 为跳过带条件格式的格）
'   2) 如需更保守的大体量确认，把 MAX_WARN 调小（默认 200,000）
Sub ToggleGreenFill_VisibleOnly()
    ' ===== 可调参数 =====
    Const GREEN_R As Long = 204
    Const GREEN_G As Long = 255
    Const GREEN_B As Long = 153
    Const MAX_WARN As Double = 200000#     ' 超过此格数弹确认
    Const skipCF As Boolean = True         ' 是否跳过带条件格式的单元格（更稳）

    ' ===== 内部变量 =====
    Dim greenColor As Long: greenColor = RGB(GREEN_R, GREEN_G, GREEN_B)
    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim sel As Range, procRange As Range, vis As Range
    Dim isSingle As Boolean
    Dim area As Range, cell As Range, subRng As Range
    Dim toGreen As Range, toClear As Range
    Dim totalCells As Double, scanned As Double
    Dim hadCF As Boolean

    ' ===== 记录并切换应用状态（Finally 一定恢复）=====
    Dim prevScr As Boolean, prevEvt As Boolean, prevDispPB As Boolean
    Dim prevCalc As XlCalculation
    With Application
        prevScr = .ScreenUpdating
        prevEvt = .EnableEvents
        prevCalc = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .StatusBar = "准备中…"
    End With
    prevDispPB = sht.DisplayPageBreaks
    sht.DisplayPageBreaks = False

    On Error GoTo Finally

    ' ===== 选择与智能缩小 =====
    Set sel = Selection
    If sel Is Nothing Then Err.Raise 5, , "没有选择区域。"

    ' 整行/整列/整表 -> 限定到 UsedRange
    If sel.Columns.Count = sht.Columns.Count Or sel.Rows.Count = sht.Rows.Count Then
        Set procRange = Intersect(sel, sht.UsedRange)
        If procRange Is Nothing Then Set procRange = sel
    Else
        Set procRange = sel
    End If

    ' 若在结构化表（ListObject）中，进一步限定到数据区（避免整个表被当作可见集）
    If Not sel.ListObject Is Nothing Then
        If Not sel.ListObject.DataBodyRange Is Nothing Then
            Dim body As Range
            Set body = sel.ListObject.DataBodyRange
            Set procRange = Intersect(procRange, body)
            If procRange Is Nothing Then Set procRange = sel
        End If
    End If

    ' ===== 单格直通，避免 SpecialCells 把范围放大 =====
    isSingle = (procRange.Areas.Count = 1 And procRange.Cells.CountLarge = 1)
    If isSingle Then
        ' 行或列被整体隐藏时视为不可见
        If procRange.EntireRow.Hidden Or procRange.EntireColumn.Hidden Then
            MsgBox "当前选中单元格不可见（所在行或列被隐藏）。", vbInformation
            GoTo Finally
        End If
        Set vis = procRange                    ' 不调用 SpecialCells
        totalCells = 1
    Else
        On Error Resume Next
        Set vis = procRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo Finally
        If vis Is Nothing Then
            MsgBox "当前选择内没有可见单元格可处理。", vbInformation
            GoTo Finally
        End If
        totalCells = vis.Cells.CountLarge
    End If

    ' ===== 大体量确认 =====
    If totalCells > MAX_WARN Then
        If MsgBox("本次将处理约 " & Format(totalCells, "#,##0") & " 个可见单元格，确认继续？", _
                  vbExclamation + vbOKCancel, "体量较大") = vbCancel Then GoTo Finally
    End If

    ' ===== 扫描并构造批量区间（合并区按代表格处理一次）=====
    scanned = 0
    For Each area In vis.Areas
        For Each cell In area.Cells
            scanned = scanned + 1
            If scanned Mod 4096 = 0 Then Application.StatusBar = _
                "扫描中… " & Format(scanned / totalCells, "0%")

            ' 处理合并区：仅以左上角代表一次；并与可见集求交
            If cell.MergeCells Then
                Set subRng = cell.MergeArea
                If cell.Address <> subRng(1, 1).Address Then GoTo ContinueCell
                Set subRng = Intersect(subRng, vis)
                If subRng Is Nothing Then GoTo ContinueCell
            Else
                Set subRng = cell
            End If

            ' 可选：跳过带条件格式的区域（常见性能/闪烁源头）
            If skipCF Then
                On Error Resume Next
                hadCF = (subRng.FormatConditions.Count > 0)
                On Error GoTo Finally
                If hadCF Then GoTo ContinueCell
            End If

            ' 判断是否当前为“绿色”
            With subRng(1, 1).Interior
                Dim isGreen As Boolean
                isGreen = (.Pattern = xlSolid And .Color = greenColor)
            End With

            If isGreen Then
                If toClear Is Nothing Then Set toClear = subRng Else Set toClear = Union(toClear, subRng)
            Else
                If toGreen Is Nothing Then Set toGreen = subRng Else Set toGreen = Union(toGreen, subRng)
            End If
ContinueCell:
        Next cell
    Next area

    ' ===== 批量写入（极大减少写操作次数）=====
    Application.StatusBar = "应用格式…"
    If Not toClear Is Nothing Then
        With toClear.Interior
            .Pattern = xlNone
            .ColorIndex = xlColorIndexNone   ' 双保险
        End With
    End If
    If Not toGreen Is Nothing Then
        With toGreen.Interior
            .Pattern = xlSolid
            .Color = greenColor
        End With
    End If

    Application.StatusBar = "完成。已处理可见单元格：" & Format(totalCells, "#,##0")

Finally:
    ' ===== 恢复应用状态 =====
    On Error Resume Next
    With Application
        .Calculation = prevCalc
        .EnableEvents = prevEvt
        .ScreenUpdating = prevScr
        .StatusBar = False
    End With
    sht.DisplayPageBreaks = prevDispPB
    On Error GoTo 0
End Sub


