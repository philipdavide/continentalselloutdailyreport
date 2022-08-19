Public Class Email
  Private Shared EMAIL_IT As String = "philip.davide@maxfinkelstein.com"
  Private Shared EMAIL_FROM As String = "noreply@maxfinkelstein.com"
  Private Shared ReadOnly SUCCESS_EMAIL_LIST() As String = {"philip.davide@maxfinkelstein.com"}
  Private Shared ReadOnly ERROR_EMAIL_LIST() As String = {"philip.davide@maxfinkelstein.com", "steve.elliot@maxfinkelstein.com"}

  Public Shared Sub LogError(Optional ByVal ex As Exception = Nothing, Optional ByVal messageType As System.Diagnostics.EventLogEntryType = Diagnostics.EventLogEntryType.Error)
    ' Log error to the Event Log

    Dim myErrorMessage As New Text.StringBuilder
    Dim myError As Exception = ex

    myErrorMessage.Append("Message" & vbCrLf & myError.Message.ToString() & vbCrLf & vbCrLf)
    myErrorMessage.Append("Source" & vbCrLf & myError.Source & vbCrLf & vbCrLf)
    If Not myError.TargetSite Is Nothing Then myErrorMessage.Append("Target site" & vbCrLf & myError.TargetSite.ToString() & vbCrLf & vbCrLf)
    myErrorMessage.Append("Stack trace" & vbCrLf & myError.StackTrace & vbCrLf & vbCrLf)
    myErrorMessage.Append("ToString()" & vbCrLf & vbCrLf & myError.ToString() & vbCrLf & vbCrLf)

    ' Assign the next InnerException
    ' to catch the details of that exception as well
    myError = myError.InnerException

    Dim msg As New System.Net.Mail.MailMessage

    msg.From = New System.Net.Mail.MailAddress(EMAIL_FROM)

    For Each email As String In ERROR_EMAIL_LIST
      msg.To.Add(email)
    Next

    msg.Subject = "An error occurred in the Web application " & My.Application.Info.AssemblyName

    If dbnet.dsn <> dbnet.MFIPROD Then
      msg.To.Clear()
      msg.To.Add(EMAIL_IT)
      msg.Subject = "TEST An error occurred in the Web application " & My.Application.Info.AssemblyName & " TEST"
    End If

    msg.Body = myErrorMessage.ToString
    'body cannot be in html otherwise the vbcrlf's will not work.
    'it'll look like one huge messy paragraph
    SendMailMessage(msg)
  End Sub

  Public Shared Sub JobCompleteEmail()
    Dim msg As New System.Net.Mail.MailMessage

    msg.From = New System.Net.Mail.MailAddress(EMAIL_FROM)

    For Each email As String In SUCCESS_EMAIL_LIST
      msg.To.Add(email)
    Next

    msg.Subject = "Continental Sellout Daily Report Complete For Date " & Now.Date.ToString("dd-MMM-yyyy")

    If dbnet.dsn <> dbnet.MFIPROD Then
      msg.To.Clear()
      msg.To.Add(EMAIL_IT)
      msg.Subject = "TEST Continental Sellout Daily Report Complete For Date " & Now.Date.ToString("dd-MMM-yyyy") & " TEST"
    End If

    SendMailMessage(msg)
  End Sub

  Public Shared Sub SendMailMessage(ByVal msg As System.Net.Mail.MailMessage)
    Dim lastexception As Exception = Nothing
    Dim SMTPSERVERS As String() = {"Smtp1.mfinetwork.org", "Smtp2.mfinetwork.org"}

    For Each svr As String In SMTPSERVERS
      Try
        Dim client As New System.Net.Mail.SmtpClient(svr)
        client.Send(msg)
        lastexception = Nothing
        Exit For
      Catch ex As Exception
        lastexception = ex
      End Try
    Next
    If Not lastexception Is Nothing Then Throw lastexception
  End Sub
End Class
