'Imports System.Xml
'Imports System.Net
'Imports System.Net.Sockets
'Imports System.Threading
'Imports System
'Imports System.Text

'Public Class cIVAConnect
'    Private ReadOnly m_MACReader As XmlTextReader
'    Private m_DeviceLicensed As Boolean = False
'    Private m_ClientType As [Enum]
'    Private ReadOnly m_ANYPORT As Integer = 0
'    Private ReadOnly m_TCPPORT As Integer = 1756
'    Private ReadOnly m_UDPPORT As Integer = 1757
'    Private ReadOnly m_UDPREPLYPORT As Int16
'    Private m_SequenceNumber As Integer
'    Private WithEvents m_UDPClient As cUDPClient

'    Public Sub New(ByVal PhysicalAddress As String)
'        Dim rand As New Random

'        'm_MACReader = New XmlTextReader("Licence.xml")
'        'Do While m_MACReader.Read
'        '    Select Case m_MACReader.NodeType
'        '        Case XmlNodeType.Text
'        '            If m_MACReader.Value = PhysicalAddress Then
'        m_DeviceLicensed = True
'        '            End If
'        '    End Select
'        'Loop
'        m_UDPREPLYPORT = rand.Next(m_UDPPORT, 65000)
'        ' Autodetect device
'        m_UDPClient = New cUDPClient
'        m_UDPClient.SendUDPbroadcast(m_UDPPORT, AutodetectRequest(m_UDPREPLYPORT))
'    End Sub

'    Public ReadOnly Property IsConnected() As Boolean
'        Get
'            If m_DeviceLicensed Then
'                IsConnected = True
'            Else
'                IsConnected = False
'            End If
'        End Get
'    End Property

'    Private Function AutodetectRequest(ByVal ReplyPort As Int16) As Byte()
'        Dim RequestPacket(11) As Byte
'        Dim SeqNum(3) As Byte
'        Dim RepPort(1) As Byte
'        Dim intIndex As Integer = 4

'        SequenceNumber = 30000

'        RequestPacket(0) = 153      ' 0x99
'        RequestPacket(1) = 57       ' 0x39
'        RequestPacket(2) = 164      ' 0xA4
'        RequestPacket(3) = 39       ' 0x27
'        ' Sequence(Number)
'        SeqNum = BitConverter.GetBytes(SequenceNumber)        ' This sequence number must be a randomized number.
'        For Each SequenceByte As Byte In SeqNum
'            RequestPacket(intIndex) = SequenceByte
'            intIndex += 1
'        Next
'        RequestPacket(8) = 0        ' 0x00
'        RequestPacket(9) = 0        ' 0x00
'        intIndex = 11
'        RepPort = BitConverter.GetBytes(ReplyPort)              ' Specifies the reply UDP port
'        For Each PortByte As Byte In RepPort
'            RequestPacket(intIndex) = PortByte
'            intIndex -= 1
'        Next
'        Return RequestPacket
'    End Function

'    Private Property SequenceNumber() As Integer
'        Get
'            Return m_SequenceNumber
'        End Get
'        Set(ByVal value As Integer)
'            Dim rand As New Random
'            m_SequenceNumber = rand.Next(value)
'        End Set
'    End Property

'    '' State object for receiving data from remote device.
'    'Private Class StateObject
'    '    ' Client socket.
'    '    Public workSocket As Socket = Nothing
'    '    ' Size of receive buffer.
'    '    Public BufferSize As Integer = 256
'    '    ' Receive buffer.
'    '    Public buffer(256) As Byte
'    '    ' Received data string.
'    '    Public sb As New StringBuilder()
'    'End Class 'StateObject

'    'Private Class AsynchronousClient
'    '    ' The port number for the remote device.
'    '    Private Shared port As Integer = 11000

'    '    ' ManualResetEvent instances signal completion.
'    '    Private Shared connectDone As New ManualResetEvent(False)
'    '    Private Shared sendDone As New ManualResetEvent(False)
'    '    Private Shared receiveDone As New ManualResetEvent(False)

'    '    ' The response from the remote device.
'    '    Private Shared response As [String] = [String].Empty


'    '    Private Shared Sub StartClient()
'    '        ' Connect to a remote device.
'    '        Try
'    '            ' Establish the remote endpoint for the socket.
'    '            ' The name of the
'    '            ' remote device is "host.contoso.com".
'    '            Dim ipHostInfo As IPHostEntry = Dns.Resolve("host.contoso.com")
'    '            Dim ipAddress As IPAddress = ipHostInfo.AddressList(0)
'    '            Dim remoteEP As New IPEndPoint(ipAddress, port)

'    '            '  Create a TCP/IP socket.
'    '            Dim client As New Socket(AddressFamily.InterNetwork, _
'    '                SocketType.Stream, ProtocolType.Tcp)

'    '            ' Connect to the remote endpoint.
'    '            client.BeginConnect(remoteEP, AddressOf ConnectCallback, client)
'    '            connectDone.WaitOne()

'    '            ' Send test data to the remote device.
'    '            Send(client, "This is a test<EOF>")
'    '            sendDone.WaitOne()

'    '            ' Receive the response from the remote device.
'    '            Receive(client)
'    '            receiveDone.WaitOne()

'    '            ' Write the response to the console.
'    '            Console.WriteLine("Response received : {0}", response)

'    '            ' Release the socket.
'    '            client.Shutdown(SocketShutdown.Both)
'    '            client.Close()

'    '        Catch e As Exception
'    '            Console.WriteLine(e.ToString())
'    '        End Try
'    '    End Sub 'StartClient


'    '    Private Shared Sub ConnectCallback(ByVal ar As IAsyncResult)
'    '        Try
'    '            ' Retrieve the socket from the state object.
'    '            Dim client As Socket = CType(ar.AsyncState, Socket)

'    '            ' Complete the connection.
'    '            client.EndConnect(ar)

'    '            Console.WriteLine("Socket connected to {0}", _
'    '                client.RemoteEndPoint.ToString())

'    '            ' Signal that the connection has been made.
'    '            connectDone.Set()
'    '        Catch e As Exception
'    '            Console.WriteLine(e.ToString())
'    '        End Try
'    '    End Sub 'ConnectCallback


'    '    Private Shared Sub Receive(ByVal client As Socket)
'    '        Try
'    '            ' Create the state object.
'    '            Dim state As New StateObject()
'    '            state.workSocket = client

'    '            ' Begin receiving the data from the remote device.
'    '            client.BeginReceive(state.buffer, 0, state.BufferSize, 0, _
'    '                AddressOf ReceiveCallback, state)
'    '        Catch e As Exception
'    '            Console.WriteLine(e.ToString())
'    '        End Try
'    '    End Sub 'Receive


'    '    Private Shared Sub ReceiveCallback(ByVal ar As IAsyncResult)
'    '        Try
'    '            ' Retrieve the state object and client socket 
'    '            ' from the asynchronous state object.
'    '            Dim state As StateObject = CType(ar.AsyncState, StateObject)
'    '            Dim client As Socket = state.workSocket

'    '            ' Read data from the remote device.
'    '            Dim bytesRead As Integer = client.EndReceive(ar)

'    '            If bytesRead > 0 Then
'    '                ' There might be more data, so store the data received so far.
'    '                state.sb.Append(Encoding.ASCII.GetString(state.buffer, 0, _
'    '                    bytesRead))

'    '                ' Get the rest of the data.
'    '                client.BeginReceive(state.buffer, 0, state.BufferSize, 0, _
'    '                    AddressOf ReceiveCallback, state)
'    '            Else
'    '                ' All the data has arrived; put it in response.
'    '                If state.sb.Length > 1 Then
'    '                    response = state.sb.ToString()
'    '                End If
'    '                ' Signal that all bytes have been received.
'    '                receiveDone.Set()
'    '            End If
'    '        Catch e As Exception
'    '            Console.WriteLine(e.ToString())
'    '        End Try
'    '    End Sub 'ReceiveCallback


'    '    Private Shared Sub Send(ByVal client As Socket, ByVal data As [String])
'    '        ' Convert the string data to byte data using ASCII encoding.
'    '        Dim byteData As Byte() = Encoding.ASCII.GetBytes(data)

'    '        ' Begin sending the data to the remote device.
'    '        client.BeginSend(byteData, 0, byteData.Length, 0, _
'    '            AddressOf SendCallback, client)
'    '    End Sub 'Send


'    '    Private Shared Sub SendCallback(ByVal ar As IAsyncResult)
'    '        Try
'    '            ' Retrieve the socket from the state object.
'    '            Dim client As Socket = CType(ar.AsyncState, Socket)

'    '            ' Complete sending the data to the remote device.
'    '            Dim bytesSent As Integer = client.EndSend(ar)
'    '            Console.WriteLine("Sent {0} bytes to server.", bytesSent)

'    '            ' Signal that all bytes have been sent.
'    '            sendDone.Set()
'    '        Catch e As Exception
'    '            Console.WriteLine(e.ToString())
'    '        End Try
'    '    End Sub 'SendCallback

'    '    'Entry point that delegates to C-style main Private Function.
'    '    Public Overloads Shared Sub Main()
'    '        System.Environment.ExitCode = _
'    '           Main(System.Environment.GetCommandLineArgs())
'    '    End Sub


'    '    Public Overloads Shared Function Main(ByVal args() As [String]) As Integer
'    '        StartClient()
'    '        Return 0
'    '    End Function 'Main
'    'End Class 'AsynchronousClient

'    Private Sub m_UDPClient_DatagramReceived(ByRef Message() As Byte) Handles m_UDPClient.DatagramReceived

'    End Sub
'End Class

