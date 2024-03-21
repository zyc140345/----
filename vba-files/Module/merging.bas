Attribute VB_Name = "merging"
Option Explicit

Public Sub scheduler()
    get_rating_result
    Application.OnTime Now + TimeValue("00:00:01"), "scheduler"
End Sub


Public Sub get_rating_result()
    ' ��ѯ�Ƿ�������δ����
    websocket.SendMessage "available"
    Dim available As String
    available = websocket.ReceiveMessage
    If available = "false" Then
        Exit Sub
    End If

    Dim merge_table As Worksheet
    Set merge_table = ThisWorkbook.Worksheets("���ܱ�")
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' ������ί����
    Dim judge_name As String
    judge_name = websocket.ReceiveMessage
    
    ' ���ղ�������
    Dim department As String
    department = websocket.ReceiveMessage
    
    ' ���� web_path
    Dim web_path As String
    web_path = websocket.ReceiveMessage
    web_path = "rating_table/" & web_path
    
    ' �������۱�����
    Dim export_dir As String
    export_dir = ThisWorkbook.path & Application.PathSeparator & judge_name
    If Not fs.FolderExists(export_dir) Then
        fs.CreateFolder export_dir
    End If
    Dim export_path As String
    export_path = export_dir & Application.PathSeparator & department & ".xlsx"
    websocket.DownloadFileHTTP web_path, export_path
    
    ' ������д����ܱ�
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
    score_row = Application.Match("�ܷ�", rating_sheet.Columns("A"), 0)
    Dim score_col As Long
    score_col = Application.Match("����������", rating_sheet.Rows("3"), 0)
    target_cell.Value = rating_sheet.Cells(score_row, score_col).Value
End Sub

Public Sub finish_merge()
    '�������λ��ƽ����
    Set cur_cell = col_begin.Offset(0, 0)
    For Each department In departments
        Dim score_range As Range
        Set score_range = merge_table.Range(merge_table.Cells(cur_cell.Row, 2), cur_cell.Offset(0, -1))
        cur_cell.Value = Application.WorksheetFunction.Average(score_range)
        Set cur_cell = cur_cell.Offset(1, 0)
    Next department
    
    '�������ܱ�
    Dim export_path As String
    export_path = root_dir & Application.PathSeparator & "���ܱ�.xlsx"
    If fs.FileExists(export_path) Then
        fs.DeleteFile export_path
    End If
    merge_workbook.SaveAs export_path
    merge_workbook.Close SaveChanges:=False
    MsgBox "���ܳɹ������ܱ��ѱ�������" & export_path
End Sub

Function get_next_col_begin(merge_table As Worksheet) As Range
    Dim col As Long
    col = merge_table.Cells(1, 1).End(xlToRight).column + 1
    Set get_next_col_begin = merge_table.Cells(1, col)
End Function

Function get_cell(merge_table As Worksheet, judge_name As String, department As String) As Range
    Dim target_cell As Range
    
    On Error Resume Next
    Set target_cell = merge_table.Rows(1).Find(judge_name).Offset(WorksheetFunction.Match(department, merge_table.Columns(2), 0) - 1)
    On Error GoTo 0
    
    Set get_cell = target_cell
End Function
