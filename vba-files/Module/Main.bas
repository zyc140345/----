Attribute VB_Name = "Main"
Option Explicit

Sub Main()
    Dim wsClient As WebSocketClient
    Dim filePath As String
    Dim uploadSuccess As Boolean
    
    ' ���� CWebSocketClient ʵ��
    Set wsClient = New WebSocketClient
    
    ' ָ��Ҫ�ϴ����ļ�·��
    filePath = Workbooks(1).Path & Application.PathSeparator & "zyc" & _
               Application.PathSeparator & "���Ͽ�ѧ�빤��ѧԺ��ί.xlsx"
    
    ' ��ʼ�� WebSocket �ͻ���
    ' ע�⣺����ķ�������ַ�Ͷ˿���Ҫ��������ʵ��������е���
    wsClient.Initialize "10.37.129.2", 8080, "/ws/140345"
    
    ' �ϴ��ļ�
    uploadSuccess = wsClient.UploadFile(filePath)
    
    If uploadSuccess Then
        MsgBox "�ļ��ϴ��ɹ���", vbInformation
    Else
        MsgBox "�ļ��ϴ�ʧ�ܡ�", vbCritical
    End If
    
    ' �ر� WebSocket ����
    wsClient.CloseConnection
End Sub
