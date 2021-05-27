Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text

' State object for receiving data from remote device.
'Friend Class StateObject
'     Client socke t.
'    Public workSocket As Socket = Nothing
'     Size of receive buffer.
'    Public BufferSize As Integer = 1024
'     Receive buffer.
'    Public buffer(BufferSize) As Byte
'     Received data string.
'    Public sb As New StringBuilder()
'End Class 'StateObject

Friend Class cTCPClient
    Private ReadOnly m_RemoteEP As IPEndPoint
    Private Shared ReadOnly m_Client As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Private Shared DeviceResponse() As Byte
    Private Shared m_Connected As Boolean
    ' ManualResetEvent instances signal completion.
    Private Shared connectDone As New ManualResetEvent(False)
    Private Shared sendDone As New ManualResetEvent(False)
    Private Shared receiveDone As New ManualResetEvent(False)

    Public Event RCPMessageReceived(ByRef Message() As Byte)

    Friend Sub New(ByVal RemoteIPAddress As IPAddress, ByVal RemotePort As Integer)
        m_RemoteEP = New IPEndPoint(RemoteIPAddress, RemotePort)
    End Sub

    Friend ReadOnly Property ClientIsConnected() As Boolean
        Get
            If Not m_Connected Then
                Connect2Client()
                m_Connected = True
            End If
            Return m_Connected
        End Get
    End Property

    Private Sub Connect2Client()

        ' Connect to device's endpoint
        m_Client.BeginConnect(m_RemoteEP, New AsyncCallback(AddressOf ConnectCallback), m_Client)

        ' Wait for connect
        connectDone.WaitOne()

    End Sub

    Private Shared Sub ConnectCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object
        Dim device As Socket = CType(ar.AsyncState, Socket)

        ' Complete the connection
        device.EndConnect(ar)

        ' Signal that the connection has been made.
        connectDone.Set()
    End Sub

    Friend Sub SendMessage(ByRef TCPData() As Byte)
        ' Send data to the remote device
        m_Client.BeginSend(TCPData, 0, TCPData.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), m_Client)
        sendDone.WaitOne()
        ' Receive the response from the remote device
        ReceiveMessages()
        receiveDone.WaitOne()
    End Sub

    Private Shared Sub SendCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object
        Dim client As Socket = CType(ar.AsyncState, Socket)

        ' Complete sending the data to the remote device
        Dim bytesSend As Integer = client.EndSend(ar)
        Console.WriteLine("Sent {0} bytes to server", bytesSend)

        ' Signal that all bytes have been sent
        sendDone.Set()
    End Sub

    'Private Shared Sub Receive()
    '    ' Create the state object
    '    Dim state As New StateObject

    '    state.workSocket = m_Client

    '    ' Begin receiving the data from the remote device.
    '    m_Client.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf _
    '        ReceiveCallback), state)
    'End Sub

    Friend Sub ReceiveMessages()
        ' Create the state object
        Dim state As New StateObject

        state.workSocket = m_Client

        ' Begin receiving messages from the remote device.
        m_Client.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf _
            ReceiveMessagesCallback), state)
    End Sub

    'Private Shared Sub ReceiveCallback(ByVal ar As IAsyncResult)
    '    ' Retrieve the state object and the client socket from the asynchronous state object
    '    Dim state As StateObject = CType(ar.AsyncState, StateObject)
    '    Dim client As Socket = state.workSocket

    '    ' Read data from from the remote device
    '    Dim bytesRead As Integer = client.EndReceive(ar)

    '    'If bytesRead > 0 Then
    '    ReDim DeviceResponse(bytesRead - 1)
    '    ' There might be more data, so store the data received so far
    '    Array.Copy(state.buffer, 0, DeviceResponse, 0, bytesRead)

    '    receiveDone.Set()
    '    'End If
    'End Sub

    Private Sub ReceiveMessagesCallback(ByVal ar As IAsyncResult)
        ' Retrieve the state object and the client socket from the asynchronous state object
        Dim state As StateObject = CType(ar.AsyncState, StateObject)
        Dim client As Socket = state.workSocket

        ' Read data from from the remote device
        Dim bytesRead As Integer = client.EndReceive(ar)

        ReDim DeviceResponse(bytesRead - 1)
        Array.Copy(state.buffer, 0, DeviceResponse, 0, bytesRead)
        receiveDone.Set()
        RaiseEvent RCPMessageReceived(DeviceResponse)

        ' Continue receiving messages from the remote device.
        m_Client.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf _
            ReceiveMessagesCallback), state)
    End Sub
End Class
