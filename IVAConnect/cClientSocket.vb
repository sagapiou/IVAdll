Imports System.Net.Sockets
Imports System.Net
Imports System.Threading

' State object for receiving data from remote device.
Friend Class StateObject
    ' Client socke t.
    Public workSocket As Socket = Nothing
    ' Size of receive buffer.
    Public BufferSize As Integer = 16384
    ' Receive buffer.
    Public buffer(BufferSize) As Byte
End Class 'StateObject

' This is an asynchronous TCP socket class for establishing non blocking communication with BOSCH devices
Friend Class cClientSocket
    Private TimeoutObject As New AutoResetEvent(False)
    Private SocketEx As Exception
    Private ReadOnly m_Socket As Socket
    Private m_DeviceIP As String
    Private MessageSent As Boolean
    Private Shared DeviceResponse() As Byte
    Private ReceiveStarted As Boolean = False
    Public Event RCPMessageReceived(ByRef Message() As Byte)
    Public Event SocketDisposedError(ByVal DeviceIP As String)


    Friend Sub New()
        ' Create the communication socket
        m_Socket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        '  m_Socket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.KeepAlive, True)
        TimeoutObject.Reset()
        SocketEx = Nothing
    End Sub

    Friend Sub CloseSocket()
        ' Release the socket.
        Try
            If Not m_Socket Is Nothing Then
                If m_Socket.Connected Then
                    m_Socket.Shutdown(SocketShutdown.Both)
                End If
                m_Socket.Close()
            End If
        Catch ex As Exception
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Error during Close socket. Error descripption is " & ex.Message, EventLogEntryType.Error)
            End If
        End Try

    End Sub

    Friend Sub Connect2DeviceAsync(ByVal DeviceIP As IPAddress, ByVal ConnectionPort As Integer)
        m_DeviceIP = DeviceIP.ToString
        ' Connect to device
        m_Socket.BeginConnect(DeviceIP, ConnectionPort, New AsyncCallback(AddressOf ConnectCallback), m_Socket)

        If TimeoutObject.WaitOne(8000, False) Then
            If Not DeviceConnected Then
                Throw SocketEx
            End If
        Else
            CloseSocket()
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Connection timeout during socket setup.Socket is closed for device " & DeviceIP.ToString, EventLogEntryType.Error)
            End If
            Throw SocketEx
        End If
    End Sub

    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        DeviceConnected = False
        Try
            Dim CallbackClient As Socket = CType(ar.AsyncState, Socket)

            CallbackClient.EndConnect(ar)
            DeviceConnected = True
            If DebuggingLevel > DebugLevels.Errors_Only Then
                DLLEventLog.WriteEntry("Connection established with device @ " & m_DeviceIP, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            DeviceConnected = False
            SocketEx = ex
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Error during setup of communication socket." & vbCrLf & ex.Message, EventLogEntryType.Error)
            End If
        Finally
            TimeoutObject.Set()
        End Try
    End Sub

    Friend Sub ReceiveMessages()
        ' Create the state object
        Dim state As New StateObject
        'DLLEventLog.WriteEntry("Beginning receive", EventLogEntryType.Warning)
        Try
            state.workSocket = m_Socket

            ' Begin receiving messages from the remote device.

            m_Socket.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf _
                ReceiveMessagesCallback), state)
        Catch Sockex As SocketException
            DLLEventLog.WriteEntry("- 159 general socket error for device: " & m_DeviceIP & " ---> " & Sockex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
        Catch ex As Exception
            DLLEventLog.WriteEntry("- 160 general socket error for device: " & m_DeviceIP & " -----> " & ex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
        End Try
    End Sub

    Private Sub ReceiveMessagesCallback(ByVal ar As IAsyncResult)
        ' Retrieve the state object and the client socket from the asynchronous state object
        Dim state As StateObject
        Dim client As Socket
        Dim bytesRead As Integer

        ' Read data from from the remote device
        Try
            state = CType(ar.AsyncState, StateObject)
            client = state.workSocket
            bytesRead = client.EndReceive(ar)
        Catch Socketex As SocketException
            DLLEventLog.WriteEntry("-150 general socket error for device: " & m_DeviceIP & " ---> " & Socketex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
            Exit Sub
        Catch ex As Exception
            DLLEventLog.WriteEntry("-151 unexpected error for Device : " & m_DeviceIP & " -----> " & ex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
            Exit Sub
        End Try

        'DLLEventLog.WriteEntry("Receive callback started", EventLogEntryType.Warning)
        Try
            If bytesRead = 0 AndAlso client.Connected = False Then
                DLLEventLog.WriteEntry("-152 Zero bytes received from IP: " & m_DeviceIP, EventLogEntryType.Error)
                RaiseEvent SocketDisposedError(m_DeviceIP)
            Else
                If bytesRead >= Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(state.buffer, 2)) Then
                    ReDim DeviceResponse(bytesRead - 1)
                    Array.Copy(state.buffer, 0, DeviceResponse, 0, bytesRead)
                    If DebuggingLevel > DebugLevels.Errors_Events Then
                        DLLEventLog.WriteEntry(bytesRead.ToString & " bytes received from device.", EventLogEntryType.Information, DeviceResponse)
                    End If
                    'ResponseReceived.Set()
                    RaiseEvent RCPMessageReceived(DeviceResponse)
                End If
            End If

        Catch ex As Exception
            DLLEventLog.WriteEntry("- 153 Unexpected error: " & m_DeviceIP & " ---> " & ex.Message, EventLogEntryType.Error)
        End Try

        ' Continue receiving messages from the remote device.
        Try
            m_Socket.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf _
                ReceiveMessagesCallback), state)
        Catch Sockex As SocketException
            DLLEventLog.WriteEntry("general socket error for device: " & m_DeviceIP & " ---> " & Sockex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
        Catch ex As Exception
            DLLEventLog.WriteEntry("general socket error for device: " & m_DeviceIP & " -----> " & ex.Message, EventLogEntryType.Error)
            RaiseEvent SocketDisposedError(m_DeviceIP)
        End Try
    End Sub

    Friend Sub SendMessage(ByRef MessageDataBytes() As Byte)
        ' Send data to the remote device

        m_Socket.BeginSend(MessageDataBytes, 0, MessageDataBytes.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), m_Socket)

        If TimeoutObject.WaitOne(8000, False) Then
            If Not MessageSent Then
                Throw SocketEx
            End If
        Else
            RaiseEvent SocketDisposedError(m_DeviceIP)
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Connection timeout during command send.Socket is closed for device " & m_DeviceIP, EventLogEntryType.Error)
            End If
        End If

    End Sub

    Private Sub SendCallback(ByVal ar As IAsyncResult)
        MessageSent = False
        Try
            Dim CallBackClient As Socket = CType(ar.AsyncState, Socket)
            Dim BytesSent As Integer

            BytesSent = CallBackClient.EndSend(ar)
            MessageSent = True
            If DebuggingLevel > DebugLevels.Errors_Events Then
                DLLEventLog.WriteEntry(BytesSent.ToString & " bytes sent to device.")
            End If
        Catch ex As Exception
            MessageSent = False
            SocketEx = ex
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Error during message sent to device." & vbCrLf & ex.Message, EventLogEntryType.Error)
            End If
        Finally
            TimeoutObject.Set()
        End Try
    End Sub

End Class
