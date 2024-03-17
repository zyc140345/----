Attribute VB_Name = "configure"

Global departments As Variant
Global cur_idx As Long

Sub begin_rating()
    ' 获取 sheet 对象
    Dim config As Worksheet
    Set config = Workbooks(1).Worksheets("配置")
    Dim rating_table As Worksheet
    Set rating_table = Workbooks(1).Worksheets("评价表")

    ' 读取单位名称
    Dim last_row As Long
    last_row = find_last_row(config.Columns("A"))
    departments = Application.Transpose(config.Range("A2:A" & last_row).Value)
    cur_idx = 1
    
    ' 弹出输入框，提示评委输入姓名
    Dim judge_name As String
    Do While True
        judge_name = InputBox("请输入您的姓名：")
        If judge_name = "" Then
            MsgBox "姓名不得为空！"
        Else
            Exit Do
        End If
    Loop
    
    ' 创建输出目录
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim export_dir As String
    export_dir = ThisWorkbook.Path & Application.PathSeparator & judge_name
    If fs.FolderExists(export_dir) Then
        fs.DeleteFolder export_dir, True
    End If
    fs.CreateFolder export_dir
    
    ' 初始化评价表
    rating_table.Range("E2").Value = "评委：" & judge_name
    rating_table.Range("A2").Value = "单位名称：" & departments(1)
    rating.clear_score
    Dim rate_next_btn As Button, rate_prev_btn As Button
    Set rate_next_btn = rating_table.Shapes("rate_next_btn").OLEFormat.Object
    Set rate_prev_btn = rating_table.Shapes("rate_prev_btn").OLEFormat.Object
    rate_next_btn.Caption = "下一个"
    rate_next_btn.Visible = True
    rate_prev_btn.Visible = False
    rating_table.Activate
End Sub

Sub merge()
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim root_dir As String
    root_dir = Workbooks(1).Path
    
    ' 读取评委姓名
    Dim judges As Object
    Set judges = fs.GetFolder(Workbooks(1).Path).SubFolders
    
    ' 读取单位名称
    Dim last_row As Long
    last_row = find_last_row(config.Columns("A"))
    departments = Application.Transpose(config.Range("A2:A" & last_row).Value)
    
    '创建并初始化汇总表
    Dim merge_workbook As Workbook
    Set merge_workbook = Workbooks.Add
    Dim merge_table As Worksheet
    Set merge_table = merge_workbook.Sheets("Sheet1")
    merge_table.Range("A1").Value = "序号"
    merge_table.Range("B1").Value = "单位名称"
    Dim cur_cell As Range
    Set cur_cell = merge_table.Range("A2")
    Dim i As Integer
    For i = 1 To UBound(departments)
        cur_cell.Value = i
        Set cur_cell = cur_cell.Offset(1, 0)
    Next i
    Set cur_cell = merge_table.Range("B2")
    For Each department In departments
        cur_cell.Value = department
        Set cur_cell = cur_cell.Offset(1, 0)
    Next department
    Set cur_cell = merge_table.Range("C1")
    For Each judge In judges
        cur_cell.Value = judge.Name
        Set cur_cell = cur_cell.Offset(0, 1)
    Next judge
    cur_cell.Value = "平均分"
    
    '将各评委对各单位的评分填入汇总表
    Dim col_begin As Range
    Set col_begin = merge_table.Range("C2")
    For Each judge In judges
        Set cur_cell = col_begin.Offset(0, 0)
        For Each department In departments
            Dim rating_workboook_path As String
            rating_workbook_path = judge.Path & Application.PathSeparator & department & ".xlsx"
            Dim rating_workbook As Workbook
            Set rating_workbook = Workbooks.Open(rating_workbook_path)
            Dim rating_sheet As Worksheet
            Set rating_sheet = rating_workbook.Worksheets(1)
            Dim score_row As Long
            score_row = Application.Match("总分", rating_sheet.Columns("A"), 0)
            score_col = Application.Match("考评组评分", rating_sheet.Rows("3"), 0)
            cur_cell.Value = rating_sheet.Cells(score_row, score_col).Value
            Set cur_cell = cur_cell.Offset(1, 0)
            rating_workbook.Close SaveChanges:=False
        Next department
        Set col_begin = col_begin.Offset(0, 1)
    Next judge
    
    '计算各单位的平均分
    Set cur_cell = col_begin.Offset(0, 0)
    For Each department In departments
        Dim score_range As Range
        Set score_range = merge_table.Range(merge_table.Cells(cur_cell.Row, 2), cur_cell.Offset(0, -1))
        cur_cell.Value = Application.WorksheetFunction.Average(score_range)
        Set cur_cell = cur_cell.Offset(1, 0)
    Next department
    
    '导出汇总表
    Dim export_path As String
    export_path = root_dir & Application.PathSeparator & "汇总表.xlsx"
    If fs.FileExists(export_path) Then
        fs.DeleteFile export_path
    End If
    merge_workbook.SaveAs export_path
    merge_workbook.Close SaveChanges:=False
    MsgBox "汇总成功！汇总表已保存至：" & export_path
End Sub

Function find_last_row(column As Range) As Long
    '查找 column 列中最后一个非空单元格索引
    find_last_row = column.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function
