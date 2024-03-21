Attribute VB_Name = "merging"
Option Explicit

Dim NextRunTime As Date

Public Sub scheduler()
    ' 轮询服务器，下载并显示已完成的评分
    get_rating_result
    NextRunTime = Now + TimeValue("00:00:01")
    Application.OnTime NextRunTime, "scheduler"
End Sub

Public Sub cancel_scheduler()
    On Error Resume Next
    Application.OnTime EarliestTime:=NextRunTime, Procedure:="scheduler", Schedule:=False
    On Error GoTo 0
End Sub


Public Sub get_rating_result()
    ' 查询是否有评分未接收
    websocket.SendMessage "available"
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    Dim available As String
    available = websocket.ReceiveMessage
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    If available = "false" Then
        Exit Sub
    End If

    Dim merge_table As Worksheet
    Set merge_table = ThisWorkbook.Worksheets("汇总表")
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' 接收评委名称
    Dim judge_name As String
    judge_name = websocket.ReceiveMessage
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    
    ' 接收部门名称
    Dim department As String
    department = websocket.ReceiveMessage
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    
    ' 接收 web_path
    Dim web_path As String
    web_path = websocket.ReceiveMessage
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    web_path = "rating_table/" + web_path
    
    ' 接收评价表数据
    Dim export_dir As String
    export_dir = ThisWorkbook.path & Application.PathSeparator & judge_name
    If Not fs.FolderExists(export_dir) Then
        fs.CreateFolder export_dir
    End If
    Dim export_path As String
    export_path = export_dir & Application.PathSeparator & department & ".xlsx"
    websocket.DownloadFileHTTP web_path, export_path
    If websocket.dwError <> ERROR_SUCCESS Then
        GoTo web_error_handle
    End If
    
    ' 将分数写入汇总表
    Dim target_cell As Range
    Set target_cell = get_cell(merge_table, judge_name, department)
    If target_cell Is Nothing Then
        Dim next_col_begin As Range
        Set next_col_begin = get_next_col_begin(merge_table)
        next_col_begin.Value = judge_name
        Set target_cell = get_cell(merge_table, judge_name, department)
    End If
    
    Dim rating_workbook As Workbook
    Set rating_workbook = Workbooks.Open(export_path)
    Dim rating_sheet As Worksheet
    Set rating_sheet = rating_workbook.Worksheets(1)
    Dim score_row As Long
    score_row = Application.Match("总分", rating_sheet.Columns("A"), 0)
    Dim score_col As Long
    score_col = Application.Match("考评组评分", rating_sheet.Rows("3"), 0)
    target_cell.Value = rating_sheet.Cells(score_row, score_col).Value
    target_cell.Interior.Color = RGB(0, 255, 0)
    rating_workbook.Close SaveChanges:=False
    Exit Sub
    
web_error_handle:
    cancel_scheduler
    websocket.CloseConnection
    MsgBox "网络异常！错误：" & websocket.dwError
End Sub

Public Sub finish_merge()
    ' 完成评分汇总

    Dim merge_table As Worksheet
    Set merge_table = ThisWorkbook.Worksheets("汇总表")
    
    ' 验证汇总
    If Not verify_merging(merge_table) Then
        MsgBox "有评委未完成评分！"
        Exit Sub
    End If

    '计算各单位的平均分
    Dim cur_cell As Range
    Set cur_cell = get_next_col_begin(merge_table)
    If cur_cell.Offset(0, -1).Value = "平均分" Then
        Set cur_cell = cur_cell.Offset(1, -1)
    Else
        cur_cell.Value = "平均分"
        Set cur_cell = cur_cell.Offset(1, 0)
    End If
    Dim department As Variant
    For Each department In departments
        Dim score_range As Range
        Set score_range = merge_table.Range(merge_table.Cells(cur_cell.Row, 2), cur_cell.Offset(0, -1))
        cur_cell.Value = Application.WorksheetFunction.Average(score_range)
        Set cur_cell = cur_cell.Offset(1, 0)
    Next department
    
    Dim export_path As String
    export_path = save_merge_table(merge_table)
    
    cancel_scheduler
    websocket.CloseConnection
    Dim finish_merge_btn As Button
    Set finish_merge_btn = merge_table.Shapes("finish_merge_btn").OLEFormat.Object
    finish_merge_btn.Visible = False
    merge_table.Cells.ClearContents
    merge_table.Cells.ClearFormats
    
    MsgBox "汇总成功！汇总表已保存至：" & export_path
    
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("配置")
    config.Activate
End Sub

Function verify_merging(merge_table As Worksheet) As Boolean
    ' 检查是否所有评委均完成评分
    Dim next_col_begin As Range
    Set next_col_begin = get_next_col_begin(merge_table)
    If next_col_begin.Offset(0, -1).Value = "单位名称" Then
        verify_merging = False
        Exit Function
    End If
    Dim left_top As Range
    Set left_top = merge_table.Cells(2, 3)
    Dim right_down As Range
    Set right_down = next_col_begin.Offset(UBound(departments) - LBound(departments) + 1, -1)
    Dim score_range As Range
    Set score_range = merge_table.Range(left_top, right_down)
    Dim valid As Boolean
    valid = True
    Dim cell As Range
    For Each cell In score_range
        If IsEmpty(cell) Or cell.Value < 0 Then
            cell.Interior.Color = RGB(255, 0, 0)
            valid = False
        Else
            cell.Interior.Color = RGB(0, 255, 0)
        End If
    Next cell
    verify_merging = valid
End Function

Function save_merge_table(merge_table) As String
    '导出汇总表
    
    ' 构造新汇总表
    Dim new_workbook As Workbook
    Set new_workbook = Workbooks.Add
    Dim new_merge_table As Worksheet
    Set new_merge_table = new_workbook.Sheets(1)
    merge_table.UsedRange.Copy
    new_merge_table.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    new_merge_table.Cells.ClearFormats
    
    '保存新评价表
    Dim export_path As String
    export_path = ThisWorkbook.path & Application.PathSeparator & "汇总表.xlsx"
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(export_path) Then
        fs.DeleteFile export_path
    End If
    new_workbook.SaveAs export_path
    new_workbook.Close SaveChanges:=False
    
    save_merge_table = export_path
End Function

Function get_next_col_begin(merge_table As Worksheet) As Range
    ' 查找下一个空列的起始单元格
    Dim col As Long
    col = merge_table.Cells(1, 1).End(xlToRight).column + 1
    Set get_next_col_begin = merge_table.Cells(1, col)
End Function

Function get_cell(merge_table As Worksheet, judge_name As String, department As String) As Range
    ' 查找指定评委和单位的分数单元格
    Dim target_cell As Range
    
    On Error Resume Next
    Set target_cell = merge_table.Rows(1).Find(judge_name).Offset(WorksheetFunction.Match(department, merge_table.Columns(2), 0) - 1)
    On Error GoTo 0
    
    Set get_cell = target_cell
End Function
