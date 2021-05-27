Imports System.Net.NetworkInformation
Imports Win32Exception = System.ComponentModel.Win32Exception
Imports System.Runtime.InteropServices.Marshal
Imports System.Runtime.InteropServices
Imports System.Net

Module NetComponents
    <Runtime.InteropServices.DllImport("iphlpapi.dll", ExactSpelling:=True)> Public Function GetIpNetTable(ByVal pIpNetTable As IntPtr, ByRef pdwSize As Int32, <MarshalAs(UnmanagedType.Bool)> ByVal bOrder As Boolean) As Int32
    End Function
    Public Const ERROR_INSUFFICIENT_BUFFER As Int32 = 122
    Public Const MAXLEN_PHYSADDR As Int32 = 8

    Dim intIndex As Integer
    Dim m_IpNetTablePtr As IntPtr
    Dim m_MacToIpDictionary = New Generic.Dictionary(Of String, String)()
    Dim m_LastMissedIpAddress = IPAddress.Broadcast.ToString()

    Friend Structure MIB_INETROW
        <MarshalAs(UnmanagedType.U4)> Public dwIndex As Int32
        <MarshalAs(UnmanagedType.U4)> Public dwPhysAddrLen As Int32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAXLEN_PHYSADDR)> Public bPhysAddr() As Byte
        <MarshalAs(UnmanagedType.U4)> Public dwAddr As UInt32
        <MarshalAs(UnmanagedType.U4)> Public dwType As Int32
    End Structure

    Friend Sub EnumerateEthernetInterfaces(ByRef IPAddress() As Byte)
        Dim adapters As NetworkInterface()
        Dim adapterProperties As IPInterfaceProperties
        Dim unicast As UnicastIPAddressInformationCollection

        adapters = NetworkInterface.GetAllNetworkInterfaces()
        For Each adapter As NetworkInterface In adapters
            If adapter.NetworkInterfaceType = NetworkInterfaceType.Ethernet AndAlso adapter.OperationalStatus = OperationalStatus.Up Then
                'ReDim Preserve IPAddress(3)
                adapterProperties = adapter.GetIPProperties()
                unicast = adapterProperties.UnicastAddresses
                For Each UniAddr In unicast
                    If Not UniAddr.Address.IsIPv6LinkLocal Then
                        IPAddress = UniAddr.Address.GetAddressBytes
                    End If
                Next
            End If
        Next adapter
    End Sub

    Public Sub PingBroadcastAddress(ByVal HostIP As String)
        Dim Octet As String()
        Dim BroadcastAddress As String
        Dim pinger As New System.Net.NetworkInformation.Ping

        Octet = Split(HostIP, ".")
        BroadcastAddress = Octet(0) & "." & Octet(1) & "." & Octet(2) & ".255"
        pinger.Send(BroadcastAddress)
    End Sub

    Public Function populateIpNetTable() As String
        Dim m_CoBytesNeeded As Int32
        Dim ReturnIP As String = String.Empty

        Try
            m_CoBytesNeeded = 0
            m_IpNetTablePtr = IntPtr.Zero
            Dim result As Int32 = GetIpNetTable(m_IpNetTablePtr, m_CoBytesNeeded, False)
            If (result <> ERROR_INSUFFICIENT_BUFFER) Then Throw New Win32Exception(result)

            m_IpNetTablePtr = AllocCoTaskMem(m_CoBytesNeeded)

            result = GetIpNetTable(m_IpNetTablePtr, m_CoBytesNeeded, False)
            If (result <> 0) Then Throw New Win32Exception(result)
            buildIpToMacDictionary()
            If m_MacToIpDictionary.trygetvalue("00-1e-22-00-00-9e", ReturnIP) Then
                Return ReturnIP
            Else
                Return "Not Found"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub buildIpToMacDictionary()

        Try
            'get table entries
            Dim entries As Int32 = ReadInt32(m_IpNetTablePtr)

            'move to first MIB_INETROW
            Dim currentbuffer As IntPtr = GetOffsetPointer(m_IpNetTablePtr, SizeOf(GetType(Int32)))

            'allocate internal table of MIB_INETROW entries
            Dim table() As MIB_INETROW = New MIB_INETROW(entries - 1) {}

            'assign entries
            For index As Int32 = 0 To entries - 1
                Dim inetrowType As System.Type = GetType(MIB_INETROW)
                Dim entryPtr As IntPtr = GetOffsetPointer(currentbuffer, _
                (index * SizeOf(inetrowType)))
                table(index) = DirectCast( _
                PtrToStructure(entryPtr, inetrowType),  _
                MIB_INETROW)
            Next

            For Each inetrow As MIB_INETROW In table
                Dim key As String = MacByteArrayToMac48String(inetrow.bPhysAddr)

                If (Not m_MacToIpDictionary.ContainsKey(key)) Then
                    Dim value As String = ConvertIPAddressToQuadString(inetrow.dwAddr)
                    m_MacToIpDictionary.Add(key, value)
                End If

            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function GetOffsetPointer(ByVal baseAddress As IntPtr, ByVal offsetSize As Long) As IntPtr
        If (offsetSize > (Long.MaxValue - baseAddress.ToInt64)) Then Throw New ArgumentOutOfRangeException("offsetSize")
        Return New IntPtr(baseAddress.ToInt64() + offsetSize)
    End Function

    Public Function ConvertIPAddressToQuadString(ByVal ip As Int64) As String
        If (ip > Convert.ToInt64(UInt32.MaxValue)) Then Throw New ArgumentOutOfRangeException("IP")

        Dim tableip As String = String.Format( _
        Globalization.CultureInfo.InvariantCulture, _
        "{0:d}.{1:d}.{2:d}.{3:d}", _
        GetOctet(ip, 0), _
        GetOctet(ip, 1), _
        GetOctet(ip, 2), _
        GetOctet(ip, 3))
        Return tableip

    End Function

    Public Function GetOctet(ByVal ip As Int64, ByVal number As Int32) As Long
        If (ip > Convert.ToInt64(UInt32.MaxValue)) Then Throw New ArgumentOutOfRangeException("IP")
        Return (ip And (&HFF << (number * 8))) >> (number * 8)
    End Function

    Public Function MacByteArrayToMac48String(ByVal macbytes As Byte()) As String

        If (macbytes.Length < 6) Then Throw New ArgumentException("Invalid array of MAC address bytes.  This application uses MAC-48 which consists of 6 address bytes.")

        Dim macTokens As String() = Array.ConvertAll(Of Byte, String)(macbytes, AddressOf ByteToHexString)
        Return (String.Join("-", macTokens).Substring(0, 17))

    End Function

    Friend Function ByteToHexString(ByVal byt As Byte) As String
        Return (byt.ToString("x2", New Globalization.NumberFormatInfo()))
    End Function

End Module