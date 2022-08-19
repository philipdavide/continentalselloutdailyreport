Imports Microsoft.VisualBasic

Public Class Logger
  Private Const LogEventSource As String = "ContinentalSelloutDailyReport"
  Private Const LogEventSourceLog As String = "Application"

  Public Shared Sub LogMessage(ByVal message As String, ByVal messageType As System.Diagnostics.EventLogEntryType)
    ' Make sure the Eventlog Exists
    Try
      If Not System.Diagnostics.EventLog.SourceExists(LogEventSource) Then
        System.Diagnostics.EventLog.CreateEventSource(LogEventSource, LogEventSourceLog)
        ' Write the entry to the event log.
        System.Diagnostics.EventLog.WriteEntry(LogEventSource, message, messageType)
      Else
        While message.Length > 32766
          message = message.Substring(32766)
        End While
        ' Write the entry to the event log.
        System.Diagnostics.EventLog.WriteEntry(LogEventSource, message, messageType)
      End If
    Catch ex As Exception
    End Try
  End Sub
End Class
