Attribute VB_Name = "Main"
Option Explicit

Sub Main()
    Dim wsClient As WebSocketClient
    Dim filePath As String
    Dim uploadSuccess As Boolean
    
    ' 创建 CWebSocketClient 实例
    Set wsClient = New WebSocketClient
    
    ' 指定要上传的文件路径
    filePath = Workbooks(1).Path & Application.PathSeparator & "zyc" & _
               Application.PathSeparator & "材料科学与工程学院纪委.xlsx"
    
    ' 初始化 WebSocket 客户端
    ' 注意：这里的服务器地址和端口需要根据您的实际情况进行调整
    wsClient.Initialize "10.37.129.2", 8080, "/ws/140345"
    
    ' 上传文件
    uploadSuccess = wsClient.UploadFile(filePath)
    
    If uploadSuccess Then
        MsgBox "文件上传成功！", vbInformation
    Else
        MsgBox "文件上传失败。", vbCritical
    End If
    
    ' 关闭 WebSocket 连接
    wsClient.CloseConnection
End Sub
