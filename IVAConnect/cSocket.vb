'*****************************************************************************************
' This class should be deleted. Use only as reference for Stavros work with UDP sockets. *
'*****************************************************************************************

Imports System.Net.Sockets
Imports System.Net

Friend Class cSocket
    Private ReadOnly m_Socket As Socket
    Private ReadOnly m_localhostIP() As Byte = {0, 0, 0, 0}
    Private m_UDPLocalEndPoint As IPEndPoint
    Private m_UDPServerEndPoint As IPEndPoint
    Private withevents m_UDPArgs As SocketAsyncEventArgs

    Public Event DatagramReceived(ByRef Message() As Byte)

    Friend Sub New(ByVal addrFamily As AddressFamily, ByVal sockType As SocketType, ByVal protoType As ProtocolType)
        If addrFamily = AddressFamily.InterNetwork OrElse addrFamily = AddressFamily.InterNetworkV6 Then
            Try
                ' Create the socket
                m_Socket = New Socket(addrFamily, sockType, protoType)

                ' Get the local IP address
                EnumerateEthernetInterfaces(m_localhostIP)
            Catch sockEx As SocketException
                DLLEventLog.WriteEntry("Socket creation failed." & vbCrLf & "Error Code : " & sockEx.SocketErrorCode & _
                                       vbCrLf & "Error Message : " & sockEx.Message, EventLogEntryType.Error)
            Catch ex As Exception
                DLLEventLog.WriteEntry("Socket creation failed. Unspecified error. Error details follow below: " & _
                                       vbCrLf & ex.Message, EventLogEntryType.Error)
            End Try
        Else
            DLLEventLog.WriteEntry("Socket creation failed. Address Family = " & addrFamily.ToString, EventLogEntryType.Error)
        End If
    End Sub

    Friend Sub BindLocalInterface(ByVal BindPort As Integer)
        Try
            Select Case m_Socket.ProtocolType
                Case ProtocolType.Tcp

                Case ProtocolType.Udp
                    m_UDPLocalEndPoint = New IPEndPoint(WhatIsMyIP, BindPort)
                    m_Socket.Bind(m_UDPLocalEndPoint)
            End Select
        Catch ex As Exception
            DLLEventLog.WriteEntry("Error occured during socket binding. Details follow : " & vbCrLf & vbTab & _
                                   ex.Message, EventLogEntryType.Error)
        End Try
    End Sub

    Private ReadOnly Property WhatIsMyIP() As IPAddress
        Get
            Dim Address As String = String.Empty

            For Each AddressByte As Byte In m_localhostIP
                Address += AddressByte.ToString
            Next
            Try
                WhatIsMyIP = IPAddress.Parse(Address)
            Catch formex As FormatException
                DLLEventLog.WriteEntry("Format exception. Provided IP address: " & Address & " is not a valid IP address.", _
                                       EventLogEntryType.Error)
                m_Socket.Close()
                Return Nothing
            Catch ex As Exception
                DLLEventLog.WriteEntry("Unspecified error occured while parsing localhost IP address. Error details follow:" & _
                                       vbCrLf & ex.Message, EventLogEntryType.Error)
                m_Socket.Close()
                Return Nothing
            End Try
        End Get
    End Property

    Friend Sub ConnectRemoteInterface(ByVal ConnectPort As Integer)

        Try
            Select Case m_Socket.ProtocolType
                Case ProtocolType.Tcp

                Case ProtocolType.Udp
                    m_UDPServerEndPoint = New IPEndPoint(GetBroadcastIP, ConnectPort)
                    m_UDPArgs = New SocketAsyncEventArgs()
                    m_UDPArgs.RemoteEndPoint = m_UDPServerEndPoint
                    m_Socket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, True)
            End Select
        Catch socketex As SocketException
            DLLEventLog.WriteEntry(socketex.Message, EventLogEntryType.Error)
        Catch ex As Exception
            DLLEventLog.WriteEntry(ex.Message, EventLogEntryType.Error)
        End Try
    End Sub

    Private Function GetBroadcastIP() As IPAddress
        Dim BroadcastIP As IPAddress
        Dim strBroadcastIP As String = String.Empty
        Dim intIndex As Integer = 0

        ' Create broadcast IP
        For Each AddressOctet As Byte In m_localhostIP
            If intIndex < 3 Then
                strBroadcastIP += AddressOctet.ToString + "."
                intIndex += 1
            End If
        Next
        strBroadcastIP &= "255"

        BroadcastIP = IPAddress.Parse(strBroadcastIP)
        Return BroadcastIP
    End Function

    Private Sub m_UDPArgs_Completed(ByVal sender As Object, ByVal e As System.Net.Sockets.SocketAsyncEventArgs) Handles m_UDPArgs.Completed
        If Not e.SocketError = SocketError.Success Then
            DLLEventLog.WriteEntry("There was a UDP socket error." & vbCrLf & "Error details: " & e.SocketError.ToString(), _
                                   EventLogEntryType.Error)
            Exit Sub
        Else
            ' Called each time a succesful socket operation completes
            Select Case e.LastOperation
                Case SocketAsyncOperation.SendTo
                    ProcessUDPPacketReceive(e)
                Case SocketAsyncOperation.ReceiveFrom
                    ' Check if succesfully received data
                    If e.BytesTransferred > 0 And e.SocketError = SocketError.Success Then
                        RaiseEvent DatagramReceived(e.Buffer)
                    End If
            End Select
        End If
    End Sub

    Friend Function SendUDP_Packet(ByRef Bytes2Send() As Byte) As Boolean
        Try
            m_UDPArgs.SetBuffer(Bytes2Send, 0, Bytes2Send.Length)
        Catch ex As Exception
            DLLEventLog.WriteEntry(ex.Message, EventLogEntryType.Error)
        End Try
        Return m_Socket.SendToAsync(m_UDPArgs)
    End Function

    Private Sub ProcessUDPPacketReceive(ByVal e As SocketAsyncEventArgs)
        'Dim token As Socket = CType(e.UserToken, Socket)

        'If token.ReceiveAsync(e) Then                    
        If m_Socket.ReceiveFromAsync(e) Then
        End If
    End Sub
End Class
