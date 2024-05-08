Attribute VB_Name = "模块1"

Sub 设置汇总统计表格式()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Set wb = Workbooks("000（全院）2024年04月校园卡充值汇总表(正式).xlsx") ' 替换为你的工作簿名称
    Set ws = wb.Sheets("2024年04月")
    
    ' 选中A8:E45单元格
    Set rng = ws.Range("A8:e45")
    
    ' 清除边框和填充
    With rng
        .Borders.LineStyle = xlNone ' 清除边框
        .Interior.ColorIndex = xlNone ' 清除填充
    End With
    
    ' 清除选区
    Set rng = Nothing
    
    
    ' 设置行高
    ws.Rows("3:5").RowHeight = 5 'line 3-5 height=5
    ws.Rows("8:47").RowHeight = 27
    
    Columns("A:F").EntireColumn.AutoFit
    '自动设置列宽
    
    ' 选中A8:F33单元格
    Set rng = ws.Range("A8:e47")
    
    ' 设置字体和字号
    With rng.Font
        .Name = "仿宋"
        .Size = 14
    End With
    
    ' 设置斑马纹底色
    For Each cell In rng.Rows
        If cell.Row Mod 2 = 1 Then
            cell.Interior.Color = RGB(230, 230, 230) '灰色
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
    
    
    '居中对齐选中单元格A列
    With rng.Columns("A")
        .HorizontalAlignment = xlCenter
    End With
    
    
    ' 居中对齐C、D列数字
    With rng.Columns("C:D")
        .HorizontalAlignment = xlCenter
    End With
    
    ' 设置E列为会计专用格式
    rng.Columns("E").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    ' 设置数字字体为Times New Roman
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Font.Name = "Times New Roman"
        End If
    Next cell
    
    ' 清除选区
    Set rng = Nothing
    
    
    ws.Rows("38").Hidden = True
    Set rng = ws.Range("7:7,40:47") '使用range属性，取得选区，注意：选择区域用逗号隔开。
    
    rng.Rows.EntireRow.Hidden = True
    '用上面语句取代下面两条语句    rng = Range("7:7,40:48")
    
    'Rows("7:7").EntireRow.Hidden = True
    
    'Rows("32:48").EntireRow.Hidden = True
    
End Sub


