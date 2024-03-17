Attribute VB_Name = "Main"
Option Explicit

Sub Main()
    Dim wsClient As WebSocketClient
    Dim filePath As String
    Dim success As Boolean
    
    ' ���� CWebSocketClient ʵ��
    Set wsClient = New WebSocketClient
    
    ' ָ��Ҫ�ϴ����ļ�·��
    filePath = Workbooks(1).Path & Application.PathSeparator & "zyc" & _
               Application.PathSeparator & "���Ͽ�ѧ�빤��ѧԺ��ί.xlsx"
    
    ' ��ʼ�� WebSocket �ͻ���
    ' ע�⣺����ķ�������ַ�Ͷ˿���Ҫ��������ʵ��������е���
    wsClient.Initialize "10.37.129.2", 8080, "/ws/140345"
    
    ' �ϴ��ļ�
    Dim message As String
    message = "Hello World!"
    success = wsClient.SendMessage(message)
    
    If success Then
        success = wsClient.UploadFile(filePath)
        success = wsClient.DownloadFile("temp.xlsx")
    Else
        MsgBox "����ʧ�ܡ�", vbCritical
    End If
    
    ' �ر� WebSocket ����
    wsClient.CloseConnection
End Sub
