Imports System
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Net
Imports System.Text
Imports System.Timers

Public Interface iIVAConnectivity
    Sub ConnectDevice(ByVal LicenseFilePath As String, ByVal DeviceIPAddress As String)
    Sub SetDebugLevel(ByVal VerboseLevel As Short)
    Sub DisconnectDevice()
    Sub PTZPanGet()
    Sub PTZPanPreposition99Get()
    'Sub DisposeMe()
    'ReadOnly Property IsConnected() As Boolean
    Event VCAAlarm(ByVal MACAddress As String, ByVal Channel As Integer, <MarshalAs(UnmanagedType.U4)> ByVal VCAAlarm As UInt32)
    Event SocketError(ByVal IPAddress As String)
    Event PanReading(ByVal IPAddress As String, ByVal vPanReading As Integer)
    Event PanPreposition99Reading(ByVal IPAddress As String, ByVal vPanReading As Integer)
    Event ConnectionStatus(ByVal IPAddress As String, ByVal ByValConnectionStatus As Boolean)
End Interface

<ComClass(cIVA.ClassId, cIVA.InterfaceId, cIVA.EventsId)> _
Public Class cIVA
    Implements iIVAConnectivity ', IDisposable
    'Private WithEvents m_UDPClient As cUDPClient
    'Private WithEvents m_TCPClient As cTCPClient
    'Private m_RequestedMACAddress As String
    'Private m_DeviceIPAddress As String
    'Private WithEvents m_UDPSocket As cSocket
    'Private m_DeviceAuthorized As Boolean

    Private Shared ResponseReceived As New ManualResetEvent(False)

    Private WithEvents Device As New cClientSocket
    Private ReadOnly m_TCPPort As Integer = 1756
    Private ReadOnly Username As String = "+service:"
    Private ReadOnly Password As String = "1+"
    Private m_DeviceConnected As Boolean = False
    Private AlarmBits As UInt32
    Private ChannelNum As Integer
    Private m_DeviceMACAddress As String
    Private m_DeviceIPAddress As String
    Private m_ClientID As Short
    Private m_SequenceNumber As Integer
    Private NumOfVideoIn As Integer
    Private RegisteredLevel As AccessLevel
    Private CPUHealth As Byte
    Private m_WatchDogTimer As System.Timers.Timer
    Private disposed As Boolean = False
    Private tmpArr As Byte() 'SAGA 

    Private Enum ReplyPacketType As Integer
        PacketNotRelevant = -1
        AutodetectDevice = 0
        VCA_Alarm = 1
        TimeoutWarning = 2
        Reply = 3
        VCA_RUNNING_STATE = 4
    End Enum

    Private Enum MessageType As Integer
        UDP_Message = 0
        TCP_Message = 1
    End Enum

    Private Enum Datatype As Byte
        ' Specifies the data type of the payload section
        FLAG = 0                ' 0x00 (1 Byte) 
        T_OCTET = 1             ' 0x01 (1 Byte)
        T_WORD = 2              ' 0x02 (2 Byte) 
        T_INT = 4               ' 0x04 (4 Byte) 
        T_DWORD = 8             ' 0x08 (4 Byte) 
        P_OCTET = 12            ' 0x0C (N Byte) 
        P_STRING = 16           ' 0x10 (N Byte) 
        P_UNICODE = 20          ' 0x14 (N Byte) 
    End Enum

    Private Enum Action As Byte
        ' Specifies the kind of the packet
        Request = 0             ' 0x00
        Reply = 1               ' 0x01
        Message = 2             ' 0x02
        Err = 3                 ' 0x03
        SpecificError = 4       ' 0x04
    End Enum

    Private Enum ClientRegistration As Byte
        Normal = 1              ' 0x01
        HookOn = 2              ' 0x02
        HookBack = 3            ' 0x03
    End Enum

    Private Enum PasswordEncryption As Byte
        PlainText = 0           ' 0x00
        MD5Hash = 1             ' 0x01
    End Enum

    Private Enum RCP_Command As Integer
        RCP_CLIENT_REGISTRATION = 65280     ' 0xFF00
        RCP_CLIENT_UNREGISTER = 65281       ' 0xFF01
        RCP_CLIENT_TIMEOUT_WARNING = 65283  ' 0xFF03
        CONF_CPU_COUNT = 2569               ' 0x0A09
        CONF_MAC_ADDRESS = 188              ' 0x00BC
        CONF_IP_STR = 124                   ' 0x007C
        CONF_NBR_OF_VIDEO_IN = 470          ' 0x01D6
        CONF_VCA_TASK_RUNNING_STATE = 2710  ' 0x0A96
        CONF_BICOM_COMMAND = 2469           ' 0x09A5
        CONF_BICOM_SRV_CONNECTED = 2585     ' 0x0A19
    End Enum

    Private Enum RCP_Reply As Integer
        RCP_CLIENT_REGISTRATION = 65280     ' 0xFF00
        RCP_CLIENT_UNREGISTER = 65281       ' 0xFF01
        RCP_CLIENT_TIMEOUT_WARNING = 65283  ' 0xFF03
        CONF_VIPROC_ALARM = 2055            ' 0x0807
        CONF_CPU_COUNT = 2569               ' 0x0A09
        CONF_MAC_ADDRESS = 188              ' 0x00BC
        CONF_IP_STR = 124                   ' 0x007C
        CONF_NBR_OF_VIDEO_IN = 470          ' 0x01D6
        VCA_TASK_RUNNING_STATE = 2710       ' 0x0A96
        CONF_BICOM_COMMAND = 2469           ' 0x09A5
        CONF_BICOM_SRV_CONNECTED = 2585     ' 0x0A19
    End Enum

    Private Enum RCP_Messages
        RCP_CLIENT_TIMEOUT_WARNING = 65283  ' 0xFF03
        CONF_VIPROC_ALARM = 2055            ' 0x0807
        CONF_VCA_TASK_RUNNING_STATE = 2710  ' 0x0A96
        CONF_RCP_TRANSFER_TRANSPARENT_DATA = 65501   '0xFFDD
    End Enum

    Private Enum RCP_Operation As Integer
        Read_Op = 0
        Write_Op = 1
    End Enum

    Private Enum RegistrationOutcome As Integer
        Failed = 0          ' 0x00
        Succesful = 1       ' 0x01
    End Enum

    Private Enum RCP_Error As Integer
        RCP_ERROR_UNKNOWN = 255             ' 0xFF
        RCP_ERROR_INVALID_VERSION = 16      ' 0x10
        RCP_ERROR_NOT_REGISTERED = 32       ' 0x20
        RCP_ERROR_INVALID_CLIENT_ID = 33    ' 0x21
        RCP_ERROR_INVALID_METHOD = 48       ' 0x30
        RCP_ERROR_INVALID_CMD = 64          ' 0x40
        RCP_ERROR_IVALID_ACCESS_TYPE = 80   ' 0x50
        RCP_ERROR_INVALID_DATA_TYPE = 96    ' 0x60
        RCP_ERROR_WRITE_ERROR = 112         ' 0x70
        RCP_ERROR_PACKET_SIZE = 128         ' 0x80
        RCP_ERROR_READ_NOT_SUPPORTED = 144  ' 0x90
        RCP_ERROR_INVALID_AUTH_LEVEL = 160  ' 0xA0
        RCP_ERROR_INVALID_SESSION_ID = 176  ' 0xB0
        RCP_ERROR_TRY_LATER = 192           ' 0xC0
    End Enum

    Private Enum AlarmFlagsMask As Integer
        MOTION = 32768                  ' Bit01 - motion flag
        GLOBAL_CHANGE = 16384           ' Bit02 - global change flag
        SIGNAL_BRIGHT = 8192            ' Bit03 - signal too bright flag
        SIGNAL_DARK = 4096              ' Bit04 - signal too dark flag
        SIGNAL_NOISY = 2048             ' Bit05 - signal too noisy flag
        IMG_BLURRY = 1024               ' Bit06 - image too blurry flag
        SIGNAL_LOSS = 512               ' Bit07 - signal loss flag
        REF_IMG_CHK = 256               ' Bit08 - reference image check failed flag
        INV_CONF_FLAG = 128             ' Bit09 - invalid configuration flag
    End Enum

    Private Enum AccessLevel As Integer
        user = 1
        service = 2
        live = 3
        unassigned = 255
    End Enum

    Private Enum BiCom_Operation As Short
        BICOM_GET = 1                           ' 0x01
        BICOM_SET = 2                           ' 0x02
        BICOM_SETGET = 3                        ' 0x03
    End Enum

    Private Enum BiComFlagID As Short
        Return_Pay_Load_Expected = 129       ' 0x81
        Return_pay_Load_With_Error = 133     ' 0x85
        Return_pay_Load_With_Error_Best_Effort = 135     ' 0x87
    End Enum

    Private Enum BiComServerID As Short
        Device_Server = 2                   ' 0x0002
        I_O_Server = 10                     ' 0x000A
        Content_Analysis_Server = 8         ' 0x0008
        Camera_Server = 4                   ' 0x0004
        PTZ_Server = 6                      ' 0x0006
    End Enum

    Private Enum BiComObjID As Short
        POS_Pan = 274                       ' 0x0112 (Radians x 10,000)
        POS_Tilt = 275                      ' 0x0113 (Radians x 10,000)
        POS_Zoom_Ticks = 276                ' 0x0114 (0 - 255 ticks)
        IDSTRING = 336                      '0x0150
        POS_PAN_PRESET_99 = 9763            ' 0x2623
    End Enum

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "733d4835-b89f-471e-b533-4e9537cd32ea"
    Public Const InterfaceId As String = "04ed4aa1-6004-49cb-a24a-23ef48fd6d3a"
    Public Const EventsId As String = "de96a7b8-21eb-48e4-856b-69472d344e05"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        DLLEventLog = New cApplicationLog("IVA")
        DebuggingLevel = My.Settings.DebugLevel
        DLLEventLog.WriteEntry("IVA Connect DLL debug level set to " & DebuggingLevel.ToString)
        ResponseReceived.Reset()
    End Sub


    Public Sub SetDebugLevel(ByVal VerboseLevel As Short) _
    Implements iIVAConnectivity.SetDebugLevel
        If DebuggingLevel <> VerboseLevel Then
            Select Case VerboseLevel
                Case DebugLevels.None
                    My.Settings.DebugLevel = DebugLevels.None
                    DebuggingLevel = DebugLevels.None
                Case DebugLevels.Errors_Only
                    My.Settings.DebugLevel = DebugLevels.Errors_Only
                    DebuggingLevel = DebugLevels.Errors_Only
                Case DebugLevels.Errors_Events
                    My.Settings.DebugLevel = DebugLevels.Errors_Events
                    DebuggingLevel = DebugLevels.Errors_Events
                Case Else
                    My.Settings.DebugLevel = DebugLevels.All
                    DebuggingLevel = DebugLevels.All
            End Select
            My.Settings.Save()
            DLLEventLog.WriteEntry("IVA Connect DLL debug level changed to " & My.Settings.DebugLevel.ToString)
        End If
    End Sub

    Public Sub DisconnectDevice() Implements iIVAConnectivity.DisconnectDevice
        ' Unregister client from device
        Try
            If m_DeviceConnected Then
                SendCommand(RCP_Command.RCP_CLIENT_UNREGISTER, RCP_Operation.Write_Op)
            End If
            m_DeviceConnected = False
            ' Dispose the timer
            If Not m_WatchDogTimer Is Nothing Then
                m_WatchDogTimer.Dispose()
            End If
        Catch ex As Exception
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Problem During Device Disconnect.", EventLogEntryType.Error)
            End If
        Finally
            ' Gracefully shut down the socket
            Device.CloseSocket()
            ' Notify the calling application about the connection status of the device
            RaiseEvent ConnectionStatus(m_DeviceIPAddress, m_DeviceConnected)
        End Try
    End Sub

    Public Sub PTZPanGet() Implements iIVAConnectivity.PTZPanGet
        SendCommand(RCP_Command.CONF_BICOM_COMMAND, RCP_Operation.Write_Op, BiComObjID.POS_Pan)
        'SendCommand(RCP_Command.CONF_BICOM_COMMAND, RCP_Operation.Read_Op)
        'SendCommand(RCP_Command.CONF_BICOM_SRV_CONNECTED, RCP_Operation.Read_Op)
    End Sub

    Public Sub PTZPreposition99PanGet() Implements iIVAConnectivity.PTZPanPreposition99Get
        SendCommand(RCP_Command.CONF_BICOM_COMMAND, RCP_Operation.Write_Op, BiComObjID.POS_PAN_PRESET_99)
        'SendCommand(RCP_Command.CONF_BICOM_COMMAND, RCP_Operation.Read_Op)
        'SendCommand(RCP_Command.CONF_BICOM_SRV_CONNECTED, RCP_Operation.Read_Op)
    End Sub


    Public Sub ConnectDevice(ByVal LicenseFilePath As String, ByVal DeviceIPAddress As String) _
Implements iIVAConnectivity.ConnectDevice
        Dim DeviceIP As IPAddress

        m_DeviceIPAddress = DeviceIPAddress
        If LicenseFilePath = String.Empty Then
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("You should provide the path for the license file.", EventLogEntryType.Warning)
            End If
            ' Notify the calling application about the connection status of the device
            RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
            'Return False
        Else
            Try
                DeviceIP = IPAddress.Parse(DeviceIPAddress)
            Catch nullex As ArgumentNullException
                If DebuggingLevel > DebugLevels.None Then
                    DLLEventLog.WriteEntry("You should provide the device IP address.", EventLogEntryType.Error)
                End If
                ' Notify the calling application about the connection status of the device
                RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
                'Return False
            Catch formatex As FormatException
                If DebuggingLevel > DebugLevels.None Then
                    DLLEventLog.WriteEntry("You should provide a valid IP address.", EventLogEntryType.Error)
                End If
                ' Notify the calling application about the connection status of the device
                RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
                'Return False
            Catch ex As Exception
                If DebuggingLevel > DebugLevels.None Then
                    DLLEventLog.WriteEntry(ex.Message, EventLogEntryType.Error)
                End If
                ' Notify the calling application about the connection status of the device
                RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
                'Return False
            End Try                         ' Retreive authorized devices
            Try

                If ValidateLicenceFile(LicenseFilePath) Then
                    ' Setup communication with device
                    ' Create the WatchDog Timer with interval of 12 minutes (720.000 milliseconds)
                    m_WatchDogTimer = New System.Timers.Timer(720000)
                    ' Hook up the timer handle
                    AddHandler m_WatchDogTimer.Elapsed, AddressOf WatchDogTimerExpired
                    ' Only raise the event the first time Interval elapses.
                    m_WatchDogTimer.AutoReset = True
                    ' Start the timer
                    m_WatchDogTimer.Enabled = True
                    ' If the timer is declared in a long-running method, use
                    ' KeepAlive to prevent garbage collection from occurring
                    ' before the method ends.
                    GC.KeepAlive(m_WatchDogTimer)
                    ConnectToDevice(DeviceIP)
                    ResponseReceived.WaitOne()
                    If m_DeviceConnected Then
                        ' Register with device
                        SendCommand(RCP_Command.RCP_CLIENT_REGISTRATION, RCP_Operation.Write_Op)
                        ' Start listening on socket
                        Device.ReceiveMessages()
                        'DLLEventLog.WriteEntry("Socket listen is started", EventLogEntryType.Warning)
                        ResponseReceived.Reset()
                        ResponseReceived.WaitOne()
                        'Do Until Not m_ClientID = 0
                        'Loop
                        ' Request device's MAC address
                        SendCommand(RCP_Command.CONF_MAC_ADDRESS, RCP_Operation.Read_Op)
                        ResponseReceived.Reset()
                        ResponseReceived.WaitOne()
                        'Do Until Not m_DeviceMACAddress Is Nothing
                        'Loop
                        ' Check if device's MAC matches one from the license file
                        If DeviceMACs.Contains(m_DeviceMACAddress) Then
                            If DebuggingLevel > DebugLevels.Errors_Only Then
                                DLLEventLog.WriteEntry("Device with IP " & DeviceIPAddress & " is authorized and connected")
                            End If

                            ' Request the number of video inputs
                            SendCommand(RCP_Command.CONF_NBR_OF_VIDEO_IN, RCP_Operation.Read_Op)
                            ResponseReceived.Reset()
                            ResponseReceived.WaitOne()
                            'Do Until Not NumOfVideoIn = 0
                            'Loop
                            'DLLEventLog.WriteEntry("SUCCESS*********************")
                            ' Notify the calling application about the connection status of the device
                            RaiseEvent ConnectionStatus(m_DeviceIPAddress, m_DeviceConnected)
                            'Return True
                        Else
                            If DebuggingLevel > DebugLevels.None Then
                                DLLEventLog.WriteEntry("Device is NOT AUTHORIZED", EventLogEntryType.Error)
                            End If
                            DisconnectDevice()
                            ' Notify the calling application about the connection status of the device
                        End If
                    Else
                        ' Notify the calling application about the connection status of the device
                        RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
                        'Return False
                    End If
                Else
                    ' Notify the calling application about the connection status of the device
                    RaiseEvent ConnectionStatus(m_DeviceIPAddress, False)
                    'Return False
                End If
            Catch ex As Exception
                If DebuggingLevel > DebugLevels.None Then
                    DLLEventLog.WriteEntry("-156 unexpected error in sub ConnectDevice() : " & ex.Message, EventLogEntryType.Error)
                    DisconnectDevice()
                End If
            End Try

        End If
    End Sub

    Private Sub WatchDogTimerExpired(ByVal source As Object, ByVal e As ElapsedEventArgs)
        ' It's been more than ten minutes since the device last send kind of event
        m_DeviceConnected = False   ' hence flag the device as disconnected
        ' Notify the calling application about the connection status of the device
        If DebuggingLevel > DebugLevels.None Then
            DLLEventLog.WriteEntry("Device " & m_DeviceMACAddress & " disconnected due to 10 min being idle.", EventLogEntryType.Error)
        End If
        DisconnectDevice()
    End Sub

    Private Function ValidateLicenceFile(ByVal FilePath As String) As Boolean
        Dim SymXMLEncryption As New cXMLLicenseEncryption
        DeviceMACs = SymXMLEncryption.ReadDataFromXML(0, FilePath)
        If DeviceMACs Is Nothing Then
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("You should provide a valid license file.", EventLogEntryType.Error)
            End If
            Return False
        Else
            If DebuggingLevel > DebugLevels.Errors_Only Then
                DLLEventLog.WriteEntry(DeviceMACs.Length.ToString & " devices are authorised", EventLogEntryType.Information)
            End If
            Return True
        End If
    End Function

    Private Sub ConnectToDevice(ByVal DeviceIP As IPAddress)
        Try
            If DebuggingLevel > DebugLevels.Errors_Events Then
                DLLEventLog.WriteEntry("Attempting connection to device @ " & DeviceIP.ToString, EventLogEntryType.Information)
            End If
            Device.Connect2DeviceAsync(DeviceIP, m_TCPPort)
            m_DeviceConnected = True
        Catch TimeEx As TimeoutException
            If DebuggingLevel > DebugLevels.Errors_Only Then
                DLLEventLog.WriteEntry("Connection timeout.", EventLogEntryType.Warning)
            End If
            m_DeviceConnected = False
        Catch ex As Exception
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry("Sub ConnectToDevice" & vbCrLf & ex.Message, EventLogEntryType.Error)
            End If
            m_DeviceConnected = False
        Finally
            ResponseReceived.Set()
        End Try
    End Sub

    Private Function PayloadArgs(ByVal Selection As RCP_Command, Optional ByVal BiComObjectID As Integer = 0) As Byte()
        Dim Parameters() As Byte
        Dim Byte4Value() As Byte = {0, 0, 0, 0}
        Dim intIndex As Integer = 0
        Dim u8 As Encoding = Encoding.UTF8

        Select Case Selection
            Case RCP_Command.RCP_CLIENT_REGISTRATION
                intIndex = 5 + Username.Length + Password.Length
                ReDim Parameters(intIndex)
                ' CONF_RCP_CLIENT_TIMEOUT_WARNING                   0xFF03
                If BitConverter.IsLittleEndian Then
                    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(RCP_Messages.RCP_CLIENT_TIMEOUT_WARNING)).CopyTo(Byte4Value, 0)
                    Parameters(0) = Byte4Value(2)
                    Parameters(1) = Byte4Value(3)
                Else
                    BitConverter.GetBytes(RCP_Messages.RCP_CLIENT_TIMEOUT_WARNING).CopyTo(Parameters, 0)
                End If
                ' CONF_VIPROC_ALARM                                 0x0807
                If BitConverter.IsLittleEndian Then
                    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(RCP_Messages.CONF_VIPROC_ALARM)).CopyTo(Byte4Value, 0)
                    Parameters(2) = Byte4Value(2)
                    Parameters(3) = Byte4Value(3)
                Else
                    BitConverter.GetBytes(RCP_Messages.CONF_VIPROC_ALARM).CopyTo(Parameters, 2)
                End If
                'u8.GetBytes(Username, 0, Username.Length, Parameters, 4)
                'u8.GetBytes(Password, 0, Password.Length, Parameters, Username.Length + 4)
                ' CONF_VCA_TASK_RUNNING_STATE                                 0x0A96
                If BitConverter.IsLittleEndian Then
                    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(RCP_Messages.CONF_VCA_TASK_RUNNING_STATE)).CopyTo(Byte4Value, 0)
                    Parameters(4) = Byte4Value(2)
                    Parameters(5) = Byte4Value(3)
                Else
                    BitConverter.GetBytes(RCP_Messages.CONF_VCA_TASK_RUNNING_STATE).CopyTo(Parameters, 4)
                End If
                '' CONF_RCP_TRANSFER_TRANSPARENT_DATA                                 0xFFDD
                'If BitConverter.IsLittleEndian Then
                '    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(RCP_Messages.CONF_RCP_TRANSFER_TRANSPARENT_DATA)).CopyTo(Byte4Value, 0)
                '    Parameters(2) = Byte4Value(2)
                '    Parameters(3) = Byte4Value(3)
                'Else
                '    BitConverter.GetBytes(RCP_Messages.CONF_RCP_TRANSFER_TRANSPARENT_DATA).CopyTo(Parameters, 2)
                'End If
                u8.GetBytes(Username, 0, Username.Length, Parameters, 6)
                u8.GetBytes(Password, 0, Password.Length, Parameters, Username.Length + 6)
                'Case RCP_Command.CONF_BICOM_COMMAND

                '' Add the password string
                'u8.GetBytes(Password, 0, Password.Length - 1, Parameters, Password.Length + Username.Length + 6)
                ''Password, 0, Password.Length - 1).CopyTo(Parameters, 2)
            Case Else
                ReDim Parameters(1)
        End Select
        Return Parameters
    End Function

    Private Sub SendCommand(ByVal Command As RCP_Command, ByVal Operation As RCP_Operation, Optional ByVal BicomOperationID As BiComObjID = 0)
        Dim CommandBytes() As Byte
        Dim ParaArray() As Byte

        Select Case Command
            Case RCP_Command.RCP_CLIENT_REGISTRATION               ' 0xFF00
                ParaArray = PayloadArgs(Command)
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.P_OCTET, Operation, Action.Request), _
                                        CreatePayload(Command, ParaArray))
            Case RCP_Command.CONF_CPU_COUNT                         ' 0x0A09
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.T_DWORD, Operation, Action.Request), _
                                              Nothing)
            Case RCP_Command.CONF_MAC_ADDRESS                       ' 0x00BC
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.P_OCTET, Operation, Action.Request), _
                                                Nothing)
            Case RCP_Command.CONF_IP_STR                            ' 0x007C
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.P_STRING, Operation, Action.Request), _
                                                Nothing)
            Case RCP_Command.RCP_CLIENT_UNREGISTER                  ' 0xFF01
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.P_OCTET, Operation, Action.Request), _
                                                Nothing)
            Case RCP_Command.CONF_NBR_OF_VIDEO_IN                   ' 0x01D6 
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.T_DWORD, Operation, Action.Request), _
                                                 Nothing)
            Case RCP_Command.CONF_BICOM_COMMAND                     ' 0x09A5
                ParaArray = PayloadArgs(Command, BicomOperationID)
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.P_OCTET, Operation, Action.Request), _
                                              CreatePayload(Command, ParaArray, BicomOperationID))
            Case RCP_Command.CONF_BICOM_SRV_CONNECTED                     ' 0x09A5
                CommandBytes = PacketizePayload(CreateHeader(Command, Datatype.FLAG, Operation, Action.Request), _
                                                 Nothing)
           Case Else
                ReDim CommandBytes(1)
                CommandBytes(0) = 0
                CommandBytes(1) = 0
        End Select
        If m_DeviceConnected Then
            Try
                If DebuggingLevel > DebugLevels.Errors_Events Then
                    DLLEventLog.WriteEntry("Attempting to send the following data:", EventLogEntryType.Information, CommandBytes)
                End If
                Device.SendMessage(CommandBytes)
            Catch TimeEx As TimeoutException
                If DebuggingLevel > DebugLevels.Errors_Events Then
                    DLLEventLog.WriteEntry("Data sent timeout", EventLogEntryType.Warning)
                End If
            Catch ex As Exception
                If DebuggingLevel > DebugLevels.None Then
                    DLLEventLog.WriteEntry("Sub SendCommand" & vbCrLf & ex.Message, EventLogEntryType.Error)
                End If
            End Try
        End If
    End Sub

    Private Function CreateHeader(ByVal Tag As UShort, ByVal CmdDataType As Datatype, ByVal Operation As RCP_Operation, _
                                  ByVal CmdAction As Action, Optional ByVal NumDes As UShort = 0) As Byte()
        Dim Header(15) As Byte

        ' Occupies two bytes. Command to be processed by VideoJet
        If BitConverter.IsLittleEndian Then
            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(CUShort(Tag))).CopyTo(Header, 0)
        Else
            BitConverter.GetBytes(CUShort(Tag)).CopyTo(Header, 0)
        End If
        ' Transfer Tag to correct position in header
        Header(0) = Header(2)
        Header(1) = Header(3)
        Header(2) = CmdDataType                         ' Data type of the payload section
        If Operation = RCP_Operation.Read_Op Then       ' High nibble is version. Version=3
            Header(3) = 48                              ' Low nibble for Read command=0x00
        Else
            Header(3) = 49                              ' Low nibble for Write command=0x01
        End If
        Header(4) = CmdAction                           ' Specify the kind of packet
        'Header(5)                                      ' Reserved byte
        ' Client ID
        If BitConverter.IsLittleEndian Then
            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(m_ClientID)).CopyTo(Header, 6)
        Else
            BitConverter.GetBytes(m_ClientID).CopyTo(Header, 6)
        End If
        'header(8)                                      ' Session ID byte 1
        'header(9)                                      ' Session ID byte 2
        'header(10)                                     ' Session ID byte 3
        'header(11)                                     ' Session ID byte 4
        ' Numeric Descriptor
        If BitConverter.IsLittleEndian Then
            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(NumDes)).CopyTo(Header, 12)
        Else
            BitConverter.GetBytes(NumDes).CopyTo(Header, 12)
        End If
        Return Header
    End Function

    Private Function CreatePayload(ByVal TagCode As RCP_Command, ByRef ParameterArray() As Byte, Optional ByVal BiComObjectID As BiComObjID = 0) As Byte()
        Dim Payload() As Byte
        Dim tmpPayLoad() As Byte = {0, 0}
        Select Case TagCode
            Case RCP_Command.RCP_CLIENT_REGISTRATION            ' 0xFF00
                ReDim Payload(7 + ParameterArray.Length)
                Payload(0) = ClientRegistration.Normal          ' Client registration type
                'Payload(1) = 0                                 ' Reserved
                If BitConverter.IsLittleEndian Then             ' Client ID
                    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(m_ClientID)).CopyTo(Payload, 2)
                Else
                    BitConverter.GetBytes(m_ClientID).CopyTo(Payload, 2)
                End If
                Payload(4) = PasswordEncryption.PlainText           ' Password and user name encryption
                Payload(5) = Username.Length + Password.Length      ' Password and user name length
                ' Number of message tags inside this packet
                If BitConverter.IsLittleEndian Then
                    BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(CShort(3))).CopyTo(Payload, 6)
                Else
                    BitConverter.GetBytes(3).CopyTo(Payload, 6)
                End If
                ' Transfer number of tags to correct position in payload array
                'Payload(6) = Payload(8)
                'Payload(7) = Payload(9)
                ParameterArray.CopyTo(Payload, 8)               ' Tag codes for the messages which should be passed to the RCP client
            Case RCP_Command.CONF_BICOM_COMMAND                 ' 0x09A5
                Select Case BiComObjectID
                    Case BiComObjID.POS_Pan
                        ' get pan value
                        ReDim Payload(4)
                        'BitConverter.GetBytes(BiComFlagID.Return_Pay_Load_Expected).CopyTo(Payload, 0)
                        If BitConverter.IsLittleEndian Then             ' Client ID
                            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(BiComServerID.PTZ_Server)).CopyTo(Payload, 0)
                        Else
                            BitConverter.GetBytes(BiComServerID.PTZ_Server).CopyTo(Payload, 0)
                        End If
                        If BitConverter.IsLittleEndian Then             ' Client ID
                            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(BiComObjID.POS_Pan)).CopyTo(Payload, 2)
                        Else
                            BitConverter.GetBytes(BiComObjID.POS_Pan).CopyTo(Payload, 2)
                        End If
                        Payload(4) = BiCom_Operation.BICOM_GET
                    Case BiComObjID.POS_PAN_PRESET_99
                        ' get pan value
                        ReDim Payload(4)
                        'BitConverter.GetBytes(BiComFlagID.Return_Pay_Load_Expected).CopyTo(Payload, 0)
                        If BitConverter.IsLittleEndian Then             ' Client ID
                            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(BiComServerID.PTZ_Server)).CopyTo(Payload, 0)
                        Else
                            BitConverter.GetBytes(BiComServerID.PTZ_Server).CopyTo(Payload, 0)
                        End If
                        If BitConverter.IsLittleEndian Then             ' Client ID
                            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(BiComObjID.POS_PAN_PRESET_99)).CopyTo(Payload, 2)
                        Else
                            BitConverter.GetBytes(BiComObjID.POS_PAN_PRESET_99).CopyTo(Payload, 2)
                        End If
                        Payload(4) = BiCom_Operation.BICOM_GET
                        'Case for other bicom commands
                    Case Else
                        ' do nothing
                End Select
            Case Else
                ReDim Payload(1)
        End Select
        Return Payload
    End Function

    Private Function PacketizePayload(ByRef Header() As Byte, ByRef Payload() As Byte) As Byte()
        Dim Package() As Byte
        Dim PacketLength As Integer
        Dim LengthBytes(1) As Byte
        Dim intIndex As Integer = 2

        If Payload Is Nothing Then
            PacketLength = Header.Length + 4
        Else
            PacketLength = Header.Length + Payload.Length + 4
        End If
        ReDim Package(PacketLength - 1)
        Package(0) = 3                                          ' TPKT Version 3
        Package(1) = 0                                          ' Reserved (should be zero)

        If BitConverter.IsLittleEndian Then
            BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(CShort(PacketLength))).CopyTo(Package, 2)
            If Not Payload Is Nothing Then
                BitConverter.GetBytes(Net.IPAddress.HostToNetworkOrder(CShort(Payload.Length))).CopyTo(Header, 14)
            End If
        Else
            BitConverter.GetBytes(CShort(PacketLength)).CopyTo(Package, 2)  ' Byte count of the complete packet (RCP packet + TPKT length
            If Not Payload Is Nothing Then
                BitConverter.GetBytes(CShort(Payload.Length)).CopyTo(Header, 14) ' Include the number of payload bytes
            End If
        End If
        Header.CopyTo(Package, 4)                               ' RCP packet header
        If Not Payload Is Nothing Then
            Payload.CopyTo(Package, Header.Length + 4)          ' RCP packet payload
        End If
        Return Package
    End Function

    Private Property SequenceNumber() As Integer
        Get
            Return m_SequenceNumber
        End Get
        Set(ByVal value As Integer)
            Dim rand As New Random
            m_SequenceNumber = rand.Next(value)
        End Set
    End Property

    'Public ReadOnly Property IsConnected() As Boolean _
    'Implements iIVAConnectivity.IsConnected
    '    Get
    '        Return m_DeviceConnected
    '    End Get
    'End Property

    Private Function AnalyzeMessageHeader(ByRef ReceivedMessage() As Byte, ByVal CurrentMessageType As MessageType) As ReplyPacketType
        Dim SeqNum(3) As Byte
        Dim PacketLength As Integer
        Dim Tag As RCP_Reply
        Dim DataTypeByte As Datatype
        Dim RCPoperation As RCP_Operation
        Dim PacketIncomplete As Boolean
        Dim PacketAction As Action
        Dim PacketClientID As Integer
        Dim PacketSessionID As Integer
        Dim PayloadLength As Integer
        Dim Payload() As Byte

        Select Case CurrentMessageType
            'Case MessageType.UDP_Message
            '    ' Check if reply to autodetect device message
            '    If ReceivedMessage(0) = 153 AndAlso ReceivedMessage(1) = 57 AndAlso ReceivedMessage(2) = 164 AndAlso ReceivedMessage(3) = 39 Then
            '        ' Verify that message received contains the correct sequence number
            '        SeqNum = BitConverter.GetBytes(Net.IPAddress.NetworkToHostOrder(SequenceNumber))
            '        If SeqNum(0) = ReceivedMessage(4) AndAlso _
            '        SeqNum(1) = ReceivedMessage(5) AndAlso _
            '        SeqNum(2) = ReceivedMessage(6) AndAlso _
            '        SeqNum(3) = ReceivedMessage(7) Then
            '            AnalyzeMessageHeader = ReplyPacketType.AutodetectDevice
            '        Else
            '            AnalyzeMessageHeader = ReplyPacketType.PacketNotRelevant
            '        End If
            '    Else
            '        AnalyzeMessageHeader = ReplyPacketType.PacketNotRelevant
            '    End If
            Case MessageType.TCP_Message
                If ReceivedMessage.Length > 0 Then
                    ' Check TPKT structure
                    If ReceivedMessage(0) = 3 And ReceivedMessage(1) = 0 Then
                        ' Get the packet length
                        PacketLength = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedMessage, 2)) ' - 4
                        ' Check if received packet is complete
                        If ReceivedMessage.Length = PacketLength Then
                            ' Determine the type of message received
                            Tag = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedMessage, 4)) And 65535 ' Bitwise operation with mask 1111111111111111
                            DataTypeByte = ReceivedMessage(6)
                            RCPoperation = ReceivedMessage(7) And 15        ' Bitwise operation with mask 1111
                            PacketIncomplete = ReceivedMessage(8) >> 7      ' Right shift 7 bits
                            PacketAction = ReceivedMessage(8) And 15        ' Bitwise operation with mask 1111
                            PacketClientID = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedMessage, 10))
                            PacketSessionID = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt32(ReceivedMessage, 12))
                            ChannelNum = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedMessage, 16))
                            PayloadLength = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedMessage, 18))
                            ReDim Payload(PayloadLength - 1)
                            Array.Copy(ReceivedMessage, 20, Payload, 0, PayloadLength)
                            ProcessPacket(Tag, PacketAction, Payload)
                            Select Case Tag
                                Case RCP_Reply.RCP_CLIENT_TIMEOUT_WARNING
                                    Return ReplyPacketType.TimeoutWarning
                                Case RCP_Reply.CONF_VIPROC_ALARM
                                    Return ReplyPacketType.VCA_Alarm
                                Case RCP_Reply.RCP_CLIENT_REGISTRATION
                                    Return ReplyPacketType.Reply
                                Case RCP_Reply.CONF_MAC_ADDRESS
                                    Return ReplyPacketType.Reply
                                Case RCP_Reply.CONF_NBR_OF_VIDEO_IN
                                    Return ReplyPacketType.Reply
                                Case RCP_Reply.RCP_CLIENT_UNREGISTER
                                    Return ReplyPacketType.Reply
                                Case RCP_Reply.CONF_CPU_COUNT
                                    Return ReplyPacketType.Reply
                                Case RCP_Reply.VCA_TASK_RUNNING_STATE
                                    Return ReplyPacketType.VCA_RUNNING_STATE
                                Case RCP_Reply.CONF_BICOM_COMMAND
                                    Return ReplyPacketType.Reply
                                Case Else
                                    If DebuggingLevel > DebugLevels.None Then
                                        DLLEventLog.WriteEntry("Received packet is not recognised", EventLogEntryType.Error)
                                    End If
                            End Select
                        Else
                            If DebuggingLevel > DebugLevels.None Then
                                DLLEventLog.WriteEntry("Received packet is incomplete.", EventLogEntryType.Error)
                            End If
                        End If
                    Else
                        AnalyzeMessageHeader = ReplyPacketType.PacketNotRelevant
                    End If
                Else
                    If DebuggingLevel > DebugLevels.Errors_Events Then
                        DLLEventLog.WriteEntry("A TCP packet was received with 0 bytes", EventLogEntryType.Warning)
                    End If
                    AnalyzeMessageHeader = ReplyPacketType.PacketNotRelevant
                End If
        End Select
    End Function

    Private Sub ProcessPacket(ByVal TagCode As RCP_Reply, ByVal PacketAction As Action, ByRef Payload As Byte())
        Select Case PacketAction
            Case Action.Reply
                'This is a reply to a request
                ProcessReplies(TagCode, Payload)
            Case Action.Err
                ' This is an error reply to a request
                ProcessErrors(Payload(0))
            Case Action.Message
                ' This is a message
                ProcessMessages(TagCode, Payload)
        End Select

    End Sub

    Private Sub ProcessReplies(ByVal TagCode As RCP_Reply, ByRef Payload As Byte())

        Select Case TagCode
            Case RCP_Reply.RCP_CLIENT_REGISTRATION            ' CONF_RCP_CLIENT_REGISTRATION reply
                If Payload(0) = RegistrationOutcome.Succesful Then
                    m_ClientID = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(Payload, 2))
                    RegisteredLevel = Payload(1)
                    If DebuggingLevel > DebugLevels.Errors_Only Then
                        DLLEventLog.WriteEntry("Succesfully registered with device." & vbCrLf & "Registration level is " & _
                                               RegisteredLevel.ToString, EventLogEntryType.Information)
                    End If
                Else
                    If DebuggingLevel > DebugLevels.None Then
                        DLLEventLog.WriteEntry("Registration with device failed.", EventLogEntryType.Error)
                    End If
                End If
            Case RCP_Reply.CONF_MAC_ADDRESS
                m_DeviceMACAddress = MacByteArrayToMac48String(Payload)
            Case RCP_Reply.RCP_CLIENT_UNREGISTER
                If Payload(0) = RegistrationOutcome.Succesful Then
                    If DebuggingLevel > DebugLevels.Errors_Only Then
                        DLLEventLog.WriteEntry("Succesfully unregistered with device.", EventLogEntryType.Information)
                    End If
                End If
            Case RCP_Reply.CONF_NBR_OF_VIDEO_IN
                NumOfVideoIn = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt32(Payload, 0))
                If DebuggingLevel > DebugLevels.Errors_Only Then
                    DLLEventLog.WriteEntry("Number of video inputs = " & NumOfVideoIn.ToString, EventLogEntryType.Information)
                End If
                '    AnalyzeMessageHeader = ReplyPacketType.Reply
                'Case RCP_Reply.RCP_CLIENT_TIMEOUT_WARNING         ' CONF_RCP_CLIENT_TIMEOUT_WARNING message
                '    AnalyzeMessageHeader = ReplyPacketType.TimeoutWarning
                'Case RCP_Reply.CONF_VIPROC_ALARM                    ' CONF_VIPROC_ALARM message
                '    AnalyzeMessageHeader = ReplyPacketType.VCA_Alarm
                'Case RCP_Reply.CONF_CPU_COUNT                       ' CONF_CPU_COUNT
                '    AnalyzeMessageHeader = ReplyPacketType.Reply
            Case RCP_Reply.CONF_BICOM_COMMAND
                Select Case Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(Payload, 0))
                    ' case for all type of bicom server commands
                    Case BiComServerID.PTZ_Server
                        ' case for all available commands per server type
                        Select Case Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(Payload, 2))
                            Case BiComObjID.POS_Pan
                                Dim vPan As Integer = 0
                                If Payload.Length >= 5 Then
                                    'For i = 5 To Payload.Length - 1
                                    '    vPan = vPan & Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(Payload, i)).ToString
                                    '    i = i + 1
                                    'Next
                                    Dim mystr As String
                                    Dim CorrectedPayload1 As String = ""
                                    Dim CorrectedPayload2 As String = ""
                                    If Payload(5) < 16 Then
                                        CorrectedPayload1 = "0" & Hex(Payload(5))
                                    Else
                                        CorrectedPayload1 = Hex(Payload(5))
                                    End If
                                    If Payload(6) < 16 Then
                                        CorrectedPayload2 = "0" & Hex(Payload(6))
                                    Else
                                        CorrectedPayload2 = Hex(Payload(6))
                                    End If

                                    mystr = CorrectedPayload1 + CorrectedPayload2

                                    'vPan = Net.IPAddress.NetworkToHostOrder(CShort(BitConverter.ToUInt16(Payload, 5)))
                                    RaiseEvent PanReading(m_DeviceIPAddress, CLng("&H" & mystr))
                                End If
                            Case BiComObjID.POS_PAN_PRESET_99
                                Dim vPan As Integer = 0
                                If Payload.Length >= 5 Then
                                    Dim mystr As String

                                    Dim CorrectedPayload1 As String = ""
                                    Dim CorrectedPayload2 As String = ""
                                    If Payload(5) < 16 Then
                                        CorrectedPayload1 = "0" & Hex(Payload(5))
                                    Else
                                        CorrectedPayload1 = Hex(Payload(5))
                                    End If
                                    If Payload(6) < 16 Then
                                        CorrectedPayload2 = "0" & Hex(Payload(6))
                                    Else
                                        CorrectedPayload2 = Hex(Payload(6))
                                    End If

                                    mystr = CorrectedPayload1 + CorrectedPayload2

                                    RaiseEvent PanPreposition99Reading(m_DeviceIPAddress, CLng("&H" & mystr))
                                End If
                            Case Else
                        End Select
                    Case Else
                End Select
        End Select

    End Sub

    Private Sub ProcessErrors(ByVal ErrorCode As RCP_Error)
        'Select Case ErrorCode
        '    Case RCP_Error.RCP_ERROR_PACKET_SIZE

        'End Select
        If DebuggingLevel > DebugLevels.None Then
            DLLEventLog.WriteEntry("Error received from device." & vbCrLf & "Error code: " & ErrorCode.ToString(), _
                                   EventLogEntryType.Error)
        End If
    End Sub

    Private Sub ProcessMessages(ByVal TagCode As RCP_Messages, ByRef Payload As Byte())
        Dim intIndex As Integer

        Select Case TagCode
            Case RCP_Messages.CONF_VIPROC_ALARM                 ' This is a VCA alarm
                Dim AlarmFlags(1) As Byte
                Dim Detector(1) As Byte
                Dim ConfigID As Byte
                For intIndex = 0 To 1
                    AlarmFlags(intIndex) = Payload(intIndex)
                    Detector(intIndex) = Payload(intIndex + 2)
                Next
                Select Case MaskVCA(AlarmFlags)
                    Case AlarmFlagsMask.MOTION
                        If DebuggingLevel > DebugLevels.Errors_Only Then
                            DLLEventLog.WriteEntry("Device MAC " & m_DeviceMACAddress & " - Video input " & ChannelNum.ToString & ". Motion alarm triggered.", _
                                                   EventLogEntryType.Information, Payload)
                        End If
                    Case AlarmFlagsMask.REF_IMG_CHK
                        If DebuggingLevel > DebugLevels.None Then
                            DLLEventLog.WriteEntry("Device MAC " & m_DeviceMACAddress & " - Video input " & ChannelNum.ToString & ". Camera tamper alarm." & _
                                                   vbCrLf & "Deviation from the reference image detected", _
                                                   EventLogEntryType.Warning, Payload)
                        End If
                    Case AlarmFlagsMask.SIGNAL_LOSS
                        If DebuggingLevel > DebugLevels.Errors_Events Then
                            DLLEventLog.WriteEntry("Device MAC " & m_DeviceMACAddress & " - Video input " & ChannelNum.ToString & ". Video loss alarm.", _
                                                   EventLogEntryType.Information, Payload)
                        End If
                End Select
                ConfigID = Payload(4)
                AlarmBits = (Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt32(Payload, 0)) >> 32) And 4294967295 ' Bit masking with 0xFFFF FFFF
            Case RCP_Messages.CONF_VCA_TASK_RUNNING_STATE           ' This is VCA running task health message
                CPUHealth = Payload(0)
        End Select
    End Sub

    Private Function MaskVCA(ByVal AlarmFlags() As Byte) As AlarmFlagsMask
        Return Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(AlarmFlags, 0)) And 65535 ' Bitmasking with 0xFFFF
    End Function

    Private Function CheckNumOfMessages(ByRef ReceivedPacket() As Byte, ByRef MessagePositions() As Integer) As Integer
        Dim intIndex As Integer = 0
        Dim PositionIndex As Integer = 0
        Dim MessageLength As Integer = 0

        Do Until intIndex >= ReceivedPacket.Length
            If ReceivedPacket(intIndex) = 3 AndAlso ReceivedPacket(intIndex + 1) = 0 Then
                MessageLength = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReceivedPacket, intIndex + 2))
                ReDim Preserve MessagePositions(PositionIndex)
                MessagePositions(PositionIndex) = intIndex
                PositionIndex += 1
                intIndex += MessageLength
            Else
                intIndex += 1
            End If
        Loop
        ' Return the number of messages inside the packet

        If ReceivedPacket.Length / PositionIndex < 17 Then
            Console.WriteLine(Now.ToString)
        End If
        Return PositionIndex
    End Function

    Private Sub m_TCPClient_RCPMessageReceived(ByRef Message1() As Byte) Handles Device.RCPMessageReceived
        Dim message As Byte()
        ' copy the byte array to a local function variable because the Message1 byte array is possible to change before the 
        ' sub finishes therefore changing its value and producing unwanted results
        Try
            message = Message1
            Dim NumberOfMessages As Integer
            Dim intIndex As Integer
            Dim MsgIndex() As Integer
            Dim MessageFragment() As Byte
            Dim FragmentLength As Integer

            ' Restart the WatchDog Timer
            m_WatchDogTimer.Enabled = False
            m_WatchDogTimer.Enabled = True

            ' Determine the number of messages include in the received packet
            Dim intIndex2 As Integer = 0
            Dim PositionIndex As Integer = 0
            Dim MessageLength As Integer = 0

            Do Until intIndex2 >= message.Length
                If message(intIndex2) = 3 AndAlso message(intIndex2 + 1) = 0 Then
                    MessageLength = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(message, intIndex2 + 2))
                    ReDim Preserve MsgIndex(PositionIndex)
                    MsgIndex(PositionIndex) = intIndex2
                    PositionIndex += 1
                    intIndex2 += MessageLength
                Else
                    intIndex2 += 1
                End If
            Loop
            NumberOfMessages = PositionIndex

            For intIndex = 0 To NumberOfMessages - 1
                FragmentLength = Net.IPAddress.NetworkToHostOrder(BitConverter.ToInt16(message, MsgIndex(intIndex) + 2))
                ReDim MessageFragment(FragmentLength - 1)
                Array.Copy(message, MsgIndex(intIndex), MessageFragment, 0, FragmentLength)
                Select Case AnalyzeMessageHeader(MessageFragment, MessageType.TCP_Message)
                    Case ReplyPacketType.VCA_Alarm
                        If Not m_DeviceMACAddress = String.Empty AndAlso NumOfVideoIn > 0 Then
                            RaiseEvent AlarmReceived(m_DeviceMACAddress, ChannelNum, AlarmBits)
                        End If
                    Case ReplyPacketType.TimeoutWarning
                        If DebuggingLevel > DebugLevels.Errors_Only Then
                            DLLEventLog.WriteEntry(m_DeviceIPAddress & "-> : Timeout warning received", EventLogEntryType.Warning)
                        End If
                        SendCommand(RCP_Command.CONF_CPU_COUNT, RCP_Operation.Read_Op)
                        If DebuggingLevel > DebugLevels.Errors_Only Then
                            DLLEventLog.WriteEntry(m_DeviceIPAddress & "-> : CPU count read command sent to device.", EventLogEntryType.Information)
                        End If
                    Case ReplyPacketType.Reply
                        'DLLEventLog.WriteEntry("About to resume thread", EventLogEntryType.Warning)
                        ResponseReceived.Set()
                    Case ReplyPacketType.VCA_RUNNING_STATE
                        If DebuggingLevel > DebugLevels.None Then
                            If CPUHealth = False Then
                                DLLEventLog.WriteEntry(m_DeviceIPAddress & "-> : VCA engine requires more CPU resources from device with MAC address " & m_DeviceMACAddress, EventLogEntryType.Error)
                            Else
                                DLLEventLog.WriteEntry(m_DeviceIPAddress & "-> : Device with MAC address " & m_DeviceMACAddress & " has enough resources for VCA engine", EventLogEntryType.Information)
                            End If
                        End If
                End Select
            Next
        Catch ex As Exception
            If DebuggingLevel > DebugLevels.None Then
                DLLEventLog.WriteEntry(m_DeviceIPAddress & "-> : -157 unexpected error during packet received for analysis. Error Description is : " & ex.Message, EventLogEntryType.Error)
            End If
        End Try
    End Sub

    Private Sub SocketErrorHandling(ByVal vIPAddress As String) Handles Device.SocketDisposedError
        RaiseEvent SocketErrorReceived(vIPAddress)
    End Sub

    Public Event AlarmReceived(ByVal MACAddress As String, ByVal Channel As Integer, <MarshalAs(UnmanagedType.U4)> _
                               ByVal AlarmBits As UInt32) Implements iIVAConnectivity.VCAAlarm
    Public Event SocketErrorReceived(ByVal IPAddress As String) Implements iIVAConnectivity.SocketError
    Public Event PanReading(ByVal IPAddress As String, ByVal vPanReading As Integer) Implements iIVAConnectivity.PanReading
    Public Event PanPreposition99Reading(ByVal IPAddress As String, ByVal vPanReading As Integer) Implements iIVAConnectivity.PanPreposition99Reading
    Public Event ConnectionStatus(ByVal IPAddress As String, ByVal ByValConnectionStatus As Boolean) _
    Implements iIVAConnectivity.ConnectionStatus
End Class


