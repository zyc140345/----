Attribute VB_Name = "rating"
Option Explicit

Sub rate_next()
    ' ������һ������
    
    Dim rating_table As Worksheet
    Set rating_table = ThisWorkbook.Worksheets("���۱�")
    Dim rate_next_btn As Button, rate_prev_btn As Button, random_rating_btn As Button
    Set rate_next_btn = rating_table.Shapes("rate_next_btn").OLEFormat.Object
    Set rate_prev_btn = rating_table.Shapes("rate_prev_btn").OLEFormat.Object
    Set random_rating_btn = rating_table.Shapes("random_rating_btn").OLEFormat.Object
    
    ' ��֤����
    Dim valid As Boolean
    valid = verify_score(rating_table)
    ' ��������
    Dim export_path As String
    export_path = save_rating_table(rating_table)
    ' �ϴ�����
    If valid Then
        Dim department As String
        department = departments(cur_idx)
        websocket.SendMessage department
        If websocket.dwError <> ERROR_SUCCESS Then
            GoTo web_error_handle
        End If
        websocket.UploadFile export_path
        If websocket.dwError <> ERROR_SUCCESS Then
            GoTo web_error_handle
        End If
    End If
    
    cur_idx = cur_idx + 1
    clear_score rating_table
    If cur_idx - 1 = LBound(departments) Then  '��ɵ�һ����λ�����ۣ���ʾ���ذ�ť
        rate_prev_btn.Visible = True
    End If
    If cur_idx < UBound(departments) Then  'δ�������һ����λ�����µ�λ���Ʋ��ָ��ѱ���ķ���
        rating_table.Range("A2").Value = "��λ���ƣ�" & departments(cur_idx)
        recover_old_scores rating_table
    ElseIf cur_idx = UBound(departments) Then  '�������һ����λ����ʾ�ύ��ť���ָ��ѱ���ķ���
        rate_next_btn.Caption = "�ύ"
        recover_old_scores rating_table
    ElseIf cur_idx > UBound(departments) Then  '������е�λ�����ۣ��������ֱ���������
        MsgBox "�����ύ�ɹ���"
        websocket.CloseConnection
        rate_next_btn.Visible = False
        rate_prev_btn.Visible = False
        random_rating_btn.Visible = False
        rating_table.Range("A2").Value = "��λ���ƣ�"
        rating_table.Range("E2").Value = "��ί��"
        Dim config As Worksheet
        Set config = ThisWorkbook.Worksheets("����")
        config.Activate
    End If
    Exit Sub
    
web_error_handle:
    websocket.CloseConnection
    MsgBox "�����쳣������" & websocket.dwError
End Sub

Sub rate_previous()
    ' ������һ������
    
    Dim rating_table As Worksheet
    Set rating_table = ThisWorkbook.Worksheets("���۱�")
    Dim rate_next_btn As Button, rate_prev_btn As Button
    Set rate_next_btn = rating_table.Shapes("rate_next_btn").OLEFormat.Object
    Set rate_prev_btn = rating_table.Shapes("rate_prev_btn").OLEFormat.Object
    
    verify_score rating_table
    save_rating_table rating_table
    cur_idx = cur_idx - 1
    recover_old_scores rating_table
    
    rating_table.Range("A2").Value = "��λ���ƣ�" & departments(cur_idx)  '���µ�λ����
    If cur_idx + 1 = UBound(departments) Then  '�����һ����λ���أ���ʾ��һ����ť
        rate_next_btn.Caption = "��һ��"
    End If
    If cur_idx = LBound(departments) Then  '���ص���һ����λ�����ط��ذ�ť
        rate_prev_btn.Visible = False
    End If
End Sub

Sub random_rating()
    ' ����������֣����ڵ���
    Dim rating_table As Worksheet
    Set rating_table = ThisWorkbook.Worksheets("���۱�")
    Dim score_range As Range
    Set score_range = get_score_range(rating_table)
    Dim cell As Range
    For Each cell In score_range
        If cell.MergeArea.Cells(1, 1).Address = cell.Address Then
            Dim total_score As Integer
            total_score = cell.Offset(0, -1).Value
            cell.Value = get_random_int(0, total_score)
        End If
    Next cell
End Sub

Function save_rating_table(rating_table As Worksheet) As String
    '���������۱�
    Dim new_workbook As Workbook
    Set new_workbook = Workbooks.Add
    Dim new_rating_table As Worksheet
    Set new_rating_table = new_workbook.Sheets(1)
    rating_table.UsedRange.Copy
    new_rating_table.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    '���������۱�
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim export_dir As String
    export_dir = Split(rating_table.Range("E2").Value, "��")(1)
    Dim export_path As String
    export_path = ThisWorkbook.path & Application.PathSeparator & export_dir & Application.PathSeparator & departments(cur_idx) & ".xlsx"
    If fs.FileExists(export_path) Then
        fs.DeleteFile export_path
    End If
    new_workbook.SaveAs export_path
    new_workbook.Close SaveChanges:=False
    
    save_rating_table = export_path
End Function

Sub recover_old_scores(rating_table As Worksheet)
    '��ԭ����
    Dim workbook_dir As String
    workbook_dir = Split(rating_table.Range("E2").Value, "��")(1)
    Dim workbook_path As String
    workbook_path = ThisWorkbook.path & Application.PathSeparator & workbook_dir & Application.PathSeparator & departments(cur_idx) & ".xlsx"

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(workbook_path) Then
        Dim old_rating_workbook As Workbook
        Set old_rating_workbook = Workbooks.Open(workbook_path)
        Dim old_rating_table As Worksheet
        Set old_rating_table = old_rating_workbook.Worksheets(1)
        
        Dim src_range As Range, dst_range As Range
        Set src_range = get_score_range(old_rating_table)
        Set dst_range = get_score_range(rating_table)
        src_range.Copy
        dst_range.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        old_rating_workbook.Close SaveChanges:=False
    End If
End Sub

Function verify_score(rating_table As Worksheet) As Boolean
    Dim score_range As Range
    Set score_range = get_score_range(rating_table)
    
    Dim valid As Boolean
    valid = True
    Dim cell As Range
    For Each cell In score_range
        If cell.MergeArea.Cells(1, 1).Address = cell.Address Then
            If IsEmpty(cell) Or cell.Value > rating_table.Cells(cell.Row, cell.column - 1).Value Then
                cell.Interior.Color = RGB(255, 0, 0)
                valid = False
            Else
                cell.Interior.Color = RGB(255, 255, 255)
            End If
        End If
    Next cell
    If Not valid Then
        MsgBox "����δ��ɵ����֣���������Ϸ���", vbExclamation + vbOKOnly, "����"
    End If
    verify_score = valid
End Function

Function clear_score(rating_table As Worksheet)
    '������۱�
    Dim score_range As Range
    Set score_range = get_score_range(rating_table)
    score_range.ClearContents
    score_range.Interior.Color = xlNone
End Function

Function get_score_range(rating_table As Worksheet) As Range
    Dim last_row As Long, col As Long
    last_row = Application.Match("�ܷ�", rating_table.Columns("A"), 0) - 1
    col = Application.Match("����������", rating_table.Rows("3"), 0)
    Set get_score_range = rating_table.Range(rating_table.Cells(4, col), rating_table.Cells(last_row, col))
End Function

Function get_random_int(min_value As Integer, max_value As Integer) As Integer
    Randomize ' ��ʼ�������������
    get_random_int = Int((max_value - min_value + 1) * Rnd + min_value)
End Function
