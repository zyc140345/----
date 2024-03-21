VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebSocketClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Constants
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_NOT_ENOUGH_MEMORY As Long = 8

' Private variables
Private hSessionHandle As LongPtr
Private hConnectionHandle As LongPtr
Private hRequestHandle As LongPtr
Private hWebSocketHandle As LongPtr
Private dwError As Long
Private bytesRead As Long
Private ServerName As String
Private path As String
Private PORT As Long

' Initialize the WebSocket client
Public Sub Initialize(Server As String, PortNumber As Long, ApiPath As String)
    Dim fStatus As Long
    Dim dwStatusCode As Long
    Dim sizeStatusCode As Long
    sizeStatusCode = 4 ' four bytes for long ' for HTTP result request
    dwError = ERROR_SUCCESS
    fStatus = False

    ServerName = Server
    PORT = PortNumber
    path = ApiPath
    
    ' Create session handle
    hSessionHandle = WinHttpOpen(StrPtr("Websocket Sample"), _
                                        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, _
                                        0, 0, 0)
    If hSessionHandle = 0 Then
        dwError = GetLastError
    End If
    
    ' Create connection handle
    hConnectionHandle = WinHttpConnect(hSessionHandle, _
                                       StrPtr(ServerName), _
                                       PORT, 0)
    If hConnectionHandle = 0 Then
        dwError = GetLastError
    End If
    
    ' Request Method
    Dim method As String
    method = "GET" ' always
    
    ' Create request handle - use 0 for null pointer to empty strings: Version, Referrer, AcceptTypes
    hRequestHandle = WinHttpOpenRequest(hConnectionHandle, StrPtr(method), StrPtr(path), 0, 0, 0, 0)
    If hRequestHandle = 0 Then
        dwError = GetLastError
    End If
    
    ' Request client protocol upgrade from http to websocket, returns true if success
    fStatus = WinHttpSetOption(hRequestHandle, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
    If (fStatus = 0) Then
        dwError = GetLastError
    End If
    
    ' Perform websocket handshake by sending the upgrade request to server
    Dim HeaderText As String
    Dim HeaderTextLength As Long
    HeaderText = ""
    HeaderText = HeaderText & "Host: " & ServerName & vbCrLf   ' may be redundant or unnecessary
    HeaderText = HeaderText & "Sec-WebSocket-Version: 13" & vbCrLf  ' 8 or 13, may be redundant or unnecessary
    HeaderText = HeaderText & "Sec-Websocket-Protocol: echo-protocol" & vbCrLf  ' subprotocol
    ' setup for API call, trim any trailing vbCrLf
    If (Right(HeaderText, 2) = vbCrLf) Then
        HeaderText = Left(HeaderText, Len(HeaderText) - 2)
    End If
    
    If Len(HeaderText) > 0 Then ' let the API figure it out
        HeaderTextLength = -1
        fStatus = WinHttpSendRequest(hRequestHandle, StrPtr(HeaderText), _
                                     HeaderTextLength, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    Else ' call without adding headers
        fStatus = WinHttpSendRequest(hRequestHandle, WINHTTP_NO_ADDITIONAL_HEADERS, _
                                     0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    End If
    If (fStatus = 0) Then
        dwError = GetLastError
        MsgBox "quitting with dwError: " & dwError
    End If
    
    ' Receive server reply
    fStatus = WinHttpReceiveResponse(hRequestHandle, 0)
    If (fStatus = 0) Then
        dwError = GetLastError
    End If
    
    ' See if the HTTP Response confirms the upgrade, with HTTP status code 101.
    fStatus = WinHttpQueryHeaders(hRequestHandle, _
                                  (WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER), _
                                  WINHTTP_HEADER_NAME_BY_INDEX, _
                                  dwStatusCode, sizeStatusCode, WINHTTP_NO_HEADER_INDEX)
    If (fStatus = 0) Then
        dwError = GetLastError
    End If
    If dwStatusCode <> 101 Then
        Debug.Print "Code needs to be 101, ending..."
        dwError = 0
    End If
    
    ' finally, get handle to websocket
    hWebSocketHandle = WinHttpWebSocketCompleteUpgrade(hRequestHandle, 0)
    If hWebSocketHandle = 0 Then
        dwError = GetLastError
    End If
    
    ' The request handle is not needed anymore. From now on we will use the websocket handle.
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
    Debug.Print "Succesfully upgraded to websocket protocol at: " & ServerName & ":" & PORT & path
End Sub

' Send a message through WebSocket
Public Function SendMessage(message As String) As Boolean
    Dim cdwMessageLength As Long
    Dim dwError As Long
    
    cdwMessageLength = 2 * Len(message)
    
    ' Send message
    dwError = WinHttpWebSocketSend(hWebSocketHandle, _
                                   WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, _
                                   StrPtr(message), cdwMessageLength)
                    
    If dwError = ERROR_SUCCESS Then
        SendMessage = True
    Else
        dwError = GetLastError
        SendMessage = False
    End If
End Function

' Receive a message through WebSocket
Public Function ReceiveMessage() As String
    Dim dwError As Long
    Dim hFile As Long
    Dim buffer() As Byte
    Dim totalBytes As Long
    
    ' Receive total bytes field
    ReDim buffer(3) ' Length field is 4 bytes
    dwError = WinHttpWebSocketReceive(hWebSocketHandle, buffer(0), 4, 0, 0)
    
    If dwError <> ERROR_SUCCESS Then
        dwError = GetLastError
        Exit Function
    End If
    
    ' Convert data length field to Long
    totalBytes = buffer(0) + buffer(1) * 2 ^ 8 + buffer(2) * 2 ^ 16 + buffer(3) * 2 ^ 24
    
    ' Receive data
    ReDim buffer(totalBytes - 1) ' Adjust buffer size as needed
    dwError = WinHttpWebSocketReceive(hWebSocketHandle, buffer(0), UBound(buffer) + 1, 0, 0)
    If dwError <> ERROR_SUCCESS Then
        dwError = GetLastError
        Exit Function
    End If
    
    ' convert buffer into string (crudely, just ignore zeros)
    ReceiveMessage = Utf8BytesToString(buffer)
End Function

' Upload a file through WebSocket
Public Function UploadFile(filePath As String) As Boolean
    Dim dwError As Long
    Dim hFile As Long
    Dim fileSize As Long
    Dim buffer() As Byte
    
    ' Open the file
    hFile = FreeFile
    Open filePath For Binary As #hFile
    fileSize = LOF(hFile)
    
    ' Read the file content into buffer
    ReDim buffer(1 To fileSize)
    Get #hFile, , buffer
    
    ' Send the file content through WebSocket
    dwError = WinHttpWebSocketSend(hWebSocketHandle, _
                                   WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, _
                                   VarPtr(buffer(1)), fileSize)
                
    If dwError = ERROR_SUCCESS Then
        UploadFile = True
    Else
        dwError = GetLastError
        UploadFile = False
    End If
    
    ' Close the file
    Close #hFile
End Function

' Download a file through WebSocket
Public Function DownloadFile(savePath As String) As Boolean
    Dim dwError As Long
    Dim hFile As Long
    Dim buffer() As Byte
    Dim totalBytes As Long
    Dim position As Long
    Dim chunkSize As Long
    
    ' Initialize chunk size
    chunkSize = 1024 ' bytes
    
    ' Open file for writing
    hFile = FreeFile
    Open savePath For Binary Access Write As #hFile
    
    ' Receive total bytes field
    ReDim buffer(3) ' Length field is 4 bytes
    dwError = WinHttpWebSocketReceive(hWebSocketHandle, buffer(0), 4, 0, 0)
    
    If dwError <> ERROR_SUCCESS Then
        dwError = GetLastError
        DownloadFile = False
        Close #hFile
        Exit Function
    End If
    
    ' Convert data length field to Long
    totalBytes = buffer(0) + buffer(1) * 2 ^ 8 + buffer(2) * 2 ^ 16 + buffer(3) * 2 ^ 24
    
    ' Receive data in chunks
    position = 0 ' Start at the beginning of the message buffer
    Do While position < totalBytes
        ' Determine the size of the next chunk
        If totalBytes - position < chunkSize Then
            chunkSize = totalBytes - position
        End If
        
        ' Resize the buffer for the chunk
        ReDim buffer(chunkSize - 1)
        
        ' Receive a chunk of data
        dwError = WinHttpWebSocketReceive(hWebSocketHandle, buffer(0), chunkSize, bytesRead, 0)
        If dwError <> ERROR_SUCCESS Then
            dwError = GetLastError
            Exit Function
        End If
        
        ' 将块直接写入文件
        Put #hFile, , buffer
        
        ' Update the position
        position = position + bytesRead
    Loop
    
    ' Close file
    Close #hFile
    DownloadFile = True
End Function

' Download a file through HTTP
Public Function DownloadFileHTTP(webPath, savePath As String) As Boolean
    Dim FileUrl As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    
    FileUrl = "http://" & SERVER_NAME & ":" & PORT & "/" & webPath
    Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpReq.Open "GET", FileUrl, False
    objXmlHttpReq.Send
    
    If objXmlHttpReq.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objXmlHttpReq.ResponseBody
        objStream.SaveToFile savePath, 2
        objStream.Close
        DownloadFileHTTP = True
    Else
        DownloadFileHTTP = False
    End If
End Function

' Close WebSocket connection
Public Sub CloseConnection()
    ' Close WebSocket handle
    If hWebSocketHandle <> 0 Then
        WinHttpCloseHandle hWebSocketHandle
        hWebSocketHandle = 0
    End If
    
    ' Close connection handle
    If hConnectionHandle <> 0 Then
        WinHttpCloseHandle hConnectionHandle
        hConnectionHandle = 0
    End If
    
    ' Close session handle
    If hSessionHandle <> 0 Then
        WinHttpCloseHandle hSessionHandle
        hSessionHandle = 0
    End If
End Sub






