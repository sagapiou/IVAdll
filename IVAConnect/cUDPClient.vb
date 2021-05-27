Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text

Friend Class cUDPClient
    ' The port number for the remote device.
    Private ReadOnly m_localhostIP() As Byte = {0, 0, 0, 0}
    Private ReadOnly m_DgramReceived() As Byte

    ' ManualResetEvent instances signal completion.
    Private Shared connectDone As New ManualResetEvent(False)
    Private Shared sendDone As New ManualResetEvent(False)
    Private Shared receiveDone As New ManualResetEvent(False)

    ' The response from the remote device.
    Private Shared response As [String] = [String].Empty

    Public Event DatagramReceived(ByRef Message() As Byte)


    Friend Sub New()
        ' Identify the localhost IP address
        EnumerateEthernetInterfaces(m_localhostIP)

        '' Connect to the remote endpoint.
        'client.BeginConnect(remoteEP, AddressOf ConnectCallback, client)
        'connectDone.WaitOne()

        '' Send test data to the remote device.
        'Send(client, "This is a test<EOF>")
        'sendDone.WaitOne()

        '' Receive the response from the remote device.
        'Receive(client)
        'receiveDone.WaitOne()

        '' Write the response to the console.
        'Console.WriteLine("Response received : {0}", response)

        '' Release the socket.
        'client.Shutdown(SocketShutdown.Both)
        'client.Close()

    End Sub

    Friend Sub SendUDPbroadcast(ByVal UDPBroadcastPort As Integer, ByVal UDPReplyPort As Integer, ByRef BroadcastMessage() As Byte)
        Dim BroadcastIP As String = String.Empty
        Dim intIndex As Integer = 0

        ' Create broadcast IP
        For Each AddressOctet As Byte In WhatsMyIP
            If intIndex < 3 Then
                BroadcastIP = BroadcastIP + AddressOctet.ToString + "."
                intIndex += 1
            End If
        Next
        BroadcastIP &= "255"

        ' Send UDP broadcast to specified port
        ' 1. Create a UDP socket.
        Dim UDPBroadcast As New Socket(AddressFamily.InterNetwork, _
            SocketType.Dgram, ProtocolType.Udp)
        ' 2. Create the remote endpoint
        Dim remoteEP As New IPEndPoint(IPAddress.Parse(BroadcastIP), UDPBroadcastPort)
        ' 3. Send the UDP broadcast synchronously
        UDPBroadcast.SendTo(BroadcastMessage, remoteEP)

        StartListener(UDPReplyPort)
    End Sub

    Private Sub StartListener(ByVal ListenPort As Integer)
        'Dim done As Boolean = False
        Dim listener As New UdpClient(ListenPort)
        Dim localEP As New IPEndPoint(New IPAddress(WhatsMyIP), ListenPort)
        Try
            'While Not done
            'Console.WriteLine("Waiting for broadcast")
            Dim bytes As Byte() = listener.Receive(localEP)
            RaiseEvent DatagramReceived(bytes)
            'Console.WriteLine("Received broadcast from {0} :", _
            'localEP.ToString())
            'Console.WriteLine( _
            '    Encoding.ASCII.GetString(bytes, 0, bytes.Length))
            'Console.WriteLine()
            'End While
        Catch e As Exception
            Console.WriteLine(e.ToString())
        Finally
            listener.Close()
        End Try
    End Sub 'StartListener

    Private ReadOnly Property WhatsMyIP() As Byte()
        Get
            Return m_localhostIP
        End Get
    End Property

End Class

'' State object for receiving data from remote device.
'Private Class StateObject
'    ' Client socke t.
'    Public workSocket As Socket = Nothing
'    ' Size of receive buffer.
'    Public BufferSize As Integer = 256
'    ' Receive buffer.
'    Public buffer(256) As Byte
'    ' Received data string.
'    Public sb As New StringBuilder()
'End Class 'StateObject

'Public Class AsynchronousClient
'    ' The port number for the remote device.
'    Private Shared port As Integer = 11000

'    ' ManualResetEvent instances signal completion.
'    Private Shared connectDone As New ManualResetEvent(False)
'    Private Shared sendDone As New ManualResetEvent(False)
'    Private Shared receiveDone As New ManualResetEvent(False)

'    ' The response from the remote device.
'    Private Shared response As [String] = [String].Empty


'    'Private Shared Sub StartClient()
'    '    ' Connect to a remote device.
'    '    Try
'    '        '' Establish the remote endpoint for the socket.
'    '        '' The name of the
'    '        '' remote device is "host.contoso.com".

'    '        'Dim ipHostInfo As IPHostEntry = Dns.Resolve("localhost")
'    '        Dim ipAddress As IPAddress = 'ipHostInfo.AddressList(0)
'    '        Dim remoteEP As New IPEndPoint(ipAddress, port)

'    '        '  Create a UDP socket.
'    '        Dim client As New Socket(AddressFamily.InterNetwork, _
'    '            SocketType.Stream, ProtocolType.Udp)

'    '        ' Connect to the remote endpoint.
'    '        client.BeginConnect(remoteEP, AddressOf ConnectCallback, client)
'    '        connectDone.WaitOne()

'    '        ' Send test data to the remote device.
'    '        Send(client, "This is a test<EOF>")
'    '        sendDone.WaitOne()

'    '        ' Receive the response from the remote device.
'    '        Receive(client)
'    '        receiveDone.WaitOne()

'    '        ' Write the response to the console.
'    '        Console.WriteLine("Response received : {0}", response)

'    '        ' Release the socket.
'    '        client.Shutdown(SocketShutdown.Both)
'    '        client.Close()

'    '    Catch e As Exception
'    '        Console.WriteLine(e.ToString())
'    '    End Try
'    'End Sub 'StartClient


'    Private Shared Sub ConnectCallback(ByVal ar As IAsyncResult)
'        Try
'            ' Retrieve the socket from the state object.
'            Dim client As Socket = CType(ar.AsyncState, Socket)

'            ' Complete the connection.
'            client.EndConnect(ar)

'            Console.WriteLine("Socket connected to {0}", _
'                client.RemoteEndPoint.ToString())

'            ' Signal that the connection has been made.
'            connectDone.Set()
'        Catch e As Exception
'            Console.WriteLine(e.ToString())
'        End Try
'    End Sub 'ConnectCallback


'    Private Shared Sub Receive(ByVal client As Socket)
'        Try
'            ' Create the state object.
'            Dim state As New StateObject()
'            state.workSocket = client

'            ' Begin receiving the data from the remote device.
'            client.BeginReceive(state.buffer, 0, state.BufferSize, 0, _
'                AddressOf ReceiveCallback, state)
'        Catch e As Exception
'            Console.WriteLine(e.ToString())
'        End Try
'    End Sub 'Receive


'    Private Shared Sub ReceiveCallback(ByVal ar As IAsyncResult)
'        Try
'            ' Retrieve the state object and client socket 
'            ' from the asynchronous state object.
'            Dim state As StateObject = CType(ar.AsyncState, StateObject)
'            Dim client As Socket = state.workSocket

'            ' Read data from the remote device.
'            Dim bytesRead As Integer = client.EndReceive(ar)

'            If bytesRead > 0 Then
'                ' There might be more data, so store the data received so far.
'                state.sb.Append(Encoding.ASCII.GetString(state.buffer, 0, _
'                    bytesRead))

'                ' Get the rest of the data.
'                client.BeginReceive(state.buffer, 0, state.BufferSize, 0, _
'                    AddressOf ReceiveCallback, state)
'            Else
'                ' All the data has arrived; put it in response.
'                If state.sb.Length > 1 Then
'                    response = state.sb.ToString()
'                End If
'                ' Signal that all bytes have been received.
'                receiveDone.Set()
'            End If
'        Catch e As Exception
'            Console.WriteLine(e.ToString())
'        End Try
'    End Sub 'ReceiveCallback


'    Private Shared Sub Send(ByVal client As Socket, ByVal data As [String])
'        ' Convert the string data to byte data using ASCII encoding.
'        Dim byteData As Byte() = Encoding.ASCII.GetBytes(data)

'        ' Begin sending the data to the remote device.
'        client.BeginSend(byteData, 0, byteData.Length, 0, _
'            AddressOf SendCallback, client)
'    End Sub 'Send


'    Private Shared Sub SendCallback(ByVal ar As IAsyncResult)
'        Try
'            ' Retrieve the socket from the state object.
'            Dim client As Socket = CType(ar.AsyncState, Socket)

'            ' Complete sending the data to the remote device.
'            Dim bytesSent As Integer = client.EndSend(ar)
'            Console.WriteLine("Sent {0} bytes to server.", bytesSent)

'            ' Signal that all bytes have been sent.
'            sendDone.Set()
'        Catch e As Exception
'            Console.WriteLine(e.ToString())
'        End Try
'    End Sub 'SendCallback

'    ''Entry point that delegates to C-style main Private Function.
'    'Public Overloads Shared Sub Main()
'    '    System.Environment.ExitCode = _
'    '       Main(System.Environment.GetCommandLineArgs())
'    'End Sub


'    'Public Overloads Shared Function Main(ByVal args() As [String]) As Integer
'    '    StartClient()
'    '    Return 0
'    'End Function 'Main
'End Class 'AsynchronousClient
