Attribute VB_Name = "rating"

Sub rate_next()
    Dim rating_table As Worksheet
    Set rating_table = Workbooks(1).Worksheets("���۱�")
    Dim rate_next_btn As Button, rate_prev_btn As Button
    Set rate_next_btn = rating_table.Shapes("rate_next_btn").OLEFormat.Object
    Set rate_prev_btn = rating_table.Shapes("rate_prev_btn").OLEFormat.Object
    
    verify_score rating_table
    save_rating_table rating_table
    
    cur_idx = cur_idx + 1
    clear_score
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
        Dim export_dir As String
        export_dir = Split(rating_table.Range("E2").Value, "��")(1)
        MsgBox "������ɣ��뽫 " & Workbooks(1).Path & Application.PathSeparator & export_dir & " ���������ֻ��ܵ��ԣ�"
        rate_next_btn.Visible = False
        rate_prev_btn.Visible = False
        rating_table.Range("A2").Value = "��λ���ƣ�"
        rating_table.Range("E2").Value = "��ί��"
    End If
End Sub

Sub rate_previous()
    Dim rating_table As Worksheet
    Set rating_table = Workbooks(1).Worksheets("���۱�")
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

Sub save_rating_table(rating_table As Worksheet)
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
    export_path = Workbooks(1).Path & Application.PathSeparator & export_dir & Application.PathSeparator & departments(cur_idx) & ".xlsx"
    If fs.FileExists(export_path) Then
        fs.DeleteFile export_path
    End If
    new_workbook.SaveAs export_path
    new_workbook.Close SaveChanges:=False
End Sub

Sub recover_old_scores(rating_table As Worksheet)
    '��ԭ����
    Dim workbook_dir As String
    workbook_dir = Split(rating_table.Range("E2").Value, "��")(1)
    Dim workbook_path As String
    workbook_path = Workbooks(1).Path & Application.PathSeparator & workbook_dir & Application.PathSeparator & departments(cur_idx) & ".xlsx"

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(workbook_path) Then
        Dim old_rating_workbook As Workbook
        Set old_rating_workbook = Workbooks.Open(workbook_path)
        Dim old_rating_table As Worksheet
        Set old_rating_table = old_rating_workbook.Worksheets(1)
        last_row = Application.Match("�ܷ�", rating_table.Columns("A"), 0) - 1
        col = Application.Match("����������", rating_table.Rows("3"), 0)
        Dim src_range As Range, dst_range As Range
        Set src_range = old_rating_table.Range(old_rating_table.Cells(4, col), old_rating_table.Cells(last_row, col))
        Set dst_range = rating_table.Range(rating_table.Cells(4, col), rating_table.Cells(last_row, col))
        src_range.Copy
        dst_range.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        old_rating_workbook.Close SaveChanges:=False
    End If
End Sub

Function verify_score(rating_table As Worksheet) As Boolean
    Dim last_row As Integer, col As Integer
    last_row = Application.Match("�ܷ�", rating_table.Columns("A"), 0) - 1
    col = Application.Match("����������", rating_table.Rows("3"), 0)
    Dim valid As Boolean
    valid = True
    For Each cell In rating_table.Range(rating_table.Cells(4, col), rating_table.Cells(last_row, col))
        If cell.MergeArea.Cells(1, 1).Address = cell.Address Then
            If IsEmpty(cell) Then
                cell.Interior.Color = RGB(255, 0, 0)
                valid = False
            Else
                cell.Interior.Color = RGB(255, 255, 255)
            End If
        End If
    Next cell
    If Not valid Then
        MsgBox "����δ��ɵ����֣�", vbExclamation + vbOKOnly, "����"
    End If
    verify_score = valid
End Function

Function clear_score()
    '������۱�
    Dim rating_table As Worksheet
    Set rating_table = Workbooks(1).Worksheets("���۱�")
    Dim last_row As Long, col As Long
    last_row = Application.Match("�ܷ�", rating_table.Columns("A"), 0) - 1
    col = Application.Match("����������", rating_table.Rows("3"), 0)
    rating_table.Range(rating_table.Cells(4, col), rating_table.Cells(last_row, col)).ClearContents
    rating_table.Range(rating_table.Cells(4, col), rating_table.Cells(last_row, col)).Interior.Color = xlNone
End Function
