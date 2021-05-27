Imports System.Diagnostics

Public Class cApplicationLog
    Dim m_EventLog As New EventLog

    Public Sub New(ByVal EventSource As String)
        ' Check if event source already exists
        If Not CheckSource("IVA-DLL") Then
            'otherwise create a new event source
            Try
                EventLog.CreateEventSource("IVA-DLL", EventSource)
            Catch ex As Exception
                MsgBox("There was an error while creating the custom Event Logger" & vbCrLf & _
                        "Error Description: " & ex.Source & vbCrLf & _
                        "Please consult VirtualControls" & vbCrLf & _
                        "Mr. Alexandros Rodis tel: +306947725409", MsgBoxStyle.Critical, EventSource & "event logger could not be created")
            End Try
        End If
        m_EventLog.Log = EventSource
        m_EventLog.Source = "IVA-DLL"
    End Sub

    Private Function CheckSource(ByVal Source As String) As Boolean
        If EventLog.SourceExists(Source) Then
            CheckSource = True
        Else
            CheckSource = False
        End If
    End Function

    Public Sub WriteEntry(ByVal LogMessage As String)
        m_EventLog.WriteEntry(LogMessage)
    End Sub

    Public Sub WriteEntry(ByVal LogMessage As String, ByVal LogType As System.Diagnostics.EventLogEntryType)
        m_EventLog.WriteEntry(LogMessage, LogType)
    End Sub

    Public Sub WriteEntry(ByVal LogMessage As String, ByVal LogType As System.Diagnostics.EventLogEntryType, ByRef LogBytes() As Byte)
        m_EventLog.WriteEntry(LogMessage, LogType, 1000, 1, LogBytes)
    End Sub
End Class
