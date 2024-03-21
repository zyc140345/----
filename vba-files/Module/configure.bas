Attribute VB_Name = "configure"
Option Explicit

Public departments As Variant
Public cur_idx As Long

Public Const SERVER_NAME = "10.11.154.117"
Public Const PORT = 5422
Public websocket As WebSocketClient


Sub begin_rating()
    ' ��ȡ sheet ����
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("����")
    Dim rating_table As Worksheet
    Set rating_table = ThisWorkbook.Worksheets("���۱�")

    ' �����������ʾ��ί��������
    Dim judge_name As String
    Do While True
        judge_name = InputBox("����������������")
        If judge_name = "" Then
            MsgBox "��������Ϊ�գ�"
        Else
            Exit Do
        End If
    Loop

    ' ��ȡ��λ����
    Dim last_row As Long
    last_row = find_last_row(config.Columns("A"))
    departments = Application.Transpose(config.Range("A2:A" & last_row).Value)
    cur_idx = 1
    
    ' �������Ŀ¼
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim export_dir As String
    export_dir = ThisWorkbook.path & Application.PathSeparator & judge_name
    If fs.FolderExists(export_dir) Then
        fs.DeleteFolder export_dir, True
    End If
    fs.CreateFolder export_dir
    
    ' ��ʼ�����۱�
    rating_table.Range("E2").Value = "��ί��" & judge_name
    rating_table.Range("A2").Value = "��λ���ƣ�" & departments(1)
    rating.clear_score
    Dim rate_next_btn As Button, rate_prev_btn As Button
    Set rate_next_btn = rating_table.Shapes("rate_next_btn").OLEFormat.Object
    Set rate_prev_btn = rating_table.Shapes("rate_prev_btn").OLEFormat.Object
    rate_next_btn.Caption = "��һ��"
    rate_next_btn.Visible = True
    rate_prev_btn.Visible = False
    
    ' ���� websocket ����
    Dim path As String
    path = "/ws/" & judge_name
    Set websocket = New WebSocketClient
    websocket.Initialize SERVER_NAME, PORT, path
    websocket.SendMessage "judge"
    
    rating_table.Activate
End Sub

Sub begin_merge()
    ' ��ȡ��λ����
    Dim last_row As Long
    last_row = find_last_row(config.Columns("A"))
    departments = Application.Transpose(config.Range("A2:A" & last_row).Value)
    
    ' ��ʼ�����ܱ�
    Dim merge_table As Worksheet
    Set merge_table = ThisWorkbook.Worksheets("���ܱ�")
    merge_table.Range("A1").Value = "���"
    merge_table.Range("B1").Value = "��λ����"
    Dim cur_cell As Range
    Set cur_cell = merge_table.Range("A2")
    Dim i As Integer
    For i = 1 To UBound(departments)
        cur_cell.Value = i
        Set cur_cell = cur_cell.Offset(1, 0)
    Next i
    Set cur_cell = merge_table.Range("B2")
    Dim department As Variant
    For Each department In departments
        cur_cell.Value = department
        Set cur_cell = cur_cell.Offset(1, 0)
    Next department
    
    ' ���� websocket ����
    Dim path As String
    path = "/ws/" & "merger"
    Set websocket = New WebSocketClient
    websocket.Initialize SERVER_NAME, PORT, path
    websocket.SendMessage "merger"
    
    scheduler
    
    merge_table.Activate
End Sub

Function find_last_row(column As Range) As Long
    '���� column �������һ���ǿյ�Ԫ������
    find_last_row = column.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function
