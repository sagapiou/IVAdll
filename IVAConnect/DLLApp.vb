Imports System.Threading

Module DLLApp
    Public DLLEventLog As cApplicationLog
    Public DeviceMACs() As String
    Public DeviceConnected As Boolean
    Public DebuggingLevel As DebugLevels


    Friend Function ByteToHexString(ByVal byt As Byte) As String
        Return (byt.ToString("x2", New Globalization.NumberFormatInfo()))
    End Function

    Public Function MacByteArrayToMac48String(ByVal macbytes As Byte()) As String

        If (macbytes.Length < 6) Then Throw New ArgumentException("Invalid array of MAC address bytes.  This application uses MAC-48 which consists of 6 address bytes.")

        Dim macTokens As String() = Array.ConvertAll(Of Byte, String)(macbytes, AddressOf ByteToHexString)
        Return (String.Join("-", macTokens).Substring(0, 17))
    End Function

    Public Function MacStringToMac48String(ByVal MAC As String) As String
        Dim FormatedMAC As String

        If (MAC.Length < 12) Then Throw New ArgumentException("Invalid MAC string.  This application uses MAC-48 which consists of 6 address bytes.")

        FormatedMAC = MAC.Insert(2, "-")
        FormatedMAC = FormatedMAC.Insert(5, "-")
        FormatedMAC = FormatedMAC.Insert(8, "-")
        FormatedMAC = FormatedMAC.Insert(11, "-")
        FormatedMAC = FormatedMAC.Insert(14, "-")
        Return FormatedMAC
    End Function

    Public Enum DebugLevels As Short
        None = 0
        Errors_Only = 1
        Errors_Events = 2
        All = 3
    End Enum
End Module
