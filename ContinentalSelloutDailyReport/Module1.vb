Imports System.Text
Imports Renci.SshNet

Module Module1
  Private isProduction As Boolean = True
  Const sftp_uploadsite_dev As String = "139.1.145.28"
  Const sftp_uploadsite_prod As String = "139.1.145.28"
  Const sftp_uploadport_dev As Integer = 9146
  Const sftp_uploadport_prod As Integer = 9346
  Const sftp_uploadlogin_dev As String = "MaFiNa1"
  Const sftp_uploadlogin_prod As String = "MaFiNa1"
  Const sftp_uploadpass_dev As String = "!CTA3mFi2na"
  Const sftp_uploadpass_prod As String = "!CTA3mFi2na"
  Const sftp_uploadpath_dev As String = "outbound"
  Const sftp_uploadpath_prod As String = "outbound"

  Sub Main()
    Console.WriteLine(My.Application.Info.AssemblyName)
    Console.WriteLine()

    Dim dataset As New DataSet
    Dim previousDate As Integer = Now.AddDays(-1).ToString("yyyyMMdd")
    Dim currentDate As Integer = Now.AddDays(-1).ToString("yyyyMMdd")

    Try
      Console.WriteLine("Pulling data.")
      dataset = dbnet.ContinentalSelloutData(previousDate, currentDate)
    Catch ex As Exception
      Logger.LogMessage(CStr(TimeOfDay) & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Stack Trace:" & vbCrLf & ex.StackTrace, EventLogEntryType.Error)

      If isProduction Then Email.LogError(ex)
      If Not isProduction Then
        Console.WriteLine(ex.Message & vbCrLf & vbCrLf & "Stack Trace:" & vbCrLf & ex.StackTrace)
        Console.ReadLine()
      End If

      Exit Sub
    End Try

    Try
      Console.WriteLine("Write to CSV file.")

      ' WriteToCSV(dataset.Tables(0))
      UploadData(dataset.Tables(0))
      If isProduction Then Email.JobCompleteEmail()
      If Not isProduction Then Console.WriteLine("Job complete")
      If Not isProduction Then Console.Read()
    Catch ex As Exception
      Logger.LogMessage(CStr(TimeOfDay) & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Stack Trace:" & vbCrLf & ex.StackTrace, EventLogEntryType.Error)

      If isProduction Then Email.LogError(ex)
      If Not isProduction Then
        Console.WriteLine(ex.Message & vbCrLf & vbCrLf & "Stack Trace:" & vbCrLf & ex.StackTrace)
        Console.ReadLine()
      End If

      Exit Sub
    End Try
  End Sub

  Private Function SoldToNumber(ByVal location As String) As String
    Dim value As String = String.Empty

    Select Case location
      Case "01"
        'astoria is also the billing location
        value = "7470705"
      Case "02"
        value = "8294608"
      Case "03"
        value = "8294614"
      Case "04"
        value = "8294615"
      Case "05"
        value = "8294616"
      Case "07"
        value = "8294617"
      Case "08"
        value = "8294618"
      Case "09"
        value = "8294619"
      Case "10"
        value = "8294620"
      Case "11"
        value = "8294621"
      Case "12"
        value = "8294622"
      Case "13"
        value = "8294623"
      Case "14"
        value = "8294624"
      Case "15"
        value = "8294625"
      Case "16"
        value = "8294626"
      Case "17"
        value = "8294627"
    End Select

    Return value.PadLeft(10, "0")
  End Function

  Private Sub WriteToCSV(ByVal datatable As DataTable)
    Dim csvstring As New StringBuilder()

    Dim totalColumns As Integer = datatable.Columns.Count - 1

    'for testing and review
    Dim columnsArray() As String = {"BillTo", "SoldTo", "CustomerLocationNo", "ArticleNo", "CustomerItemNo", "InvoiceNo", "InvoiceLineNo", "TransactionDate", "SoldQuantity", "OnHandQuantity", "Type", "Comment1", "Comment2", "Comment3"}

    For i As Integer = 0 To columnsArray.Count - 1
      Dim val As String = columnsArray(i)
      If val.Contains(",") Then val = """" & val & """"
      csvstring.Append(val)
      If i < columnsArray.Count - 1 Then csvstring.Append("|")
    Next

    csvstring.AppendLine()

    For Each row As DataRow In datatable.Rows
      For col As Integer = 0 To totalColumns
        Dim val As String = IsNull(row(col))
        If val.Contains(",") Then val = """" & val & """"
        If datatable.Columns.Item(col).ColumnName = "SALOC" Then val = SoldToNumber(val)
        csvstring.Append(val)
        If col < totalColumns Then csvstring.Append("|")
      Next

      csvstring.AppendLine()
    Next

    Dim csvFile As String = String.Format("WS_MFI_{0}_{1}.txt",
                                          Now.ToString("yyyyMMdd"),
                                          Now.ToString("HHmm"))

    System.IO.File.WriteAllText(csvFile, csvstring.ToString())
  End Sub

  Private Sub UploadData(ByVal dt As DataTable)
    Dim content As New StringBuilder

    For Each dr As DataRow In dt.Rows
      For i As Integer = 0 To dt.Columns.Count - 1
        Dim val As String = IsNull(dr(i))
        If val.Contains(",") Then val = """" & val & """"
        If dt.Columns.Item(i).ColumnName = "SALOC" Then val = SoldToNumber(val)
        content.Append(val)
        If i < dt.Columns.Count - 1 Then content.Append("|")
      Next
      content.AppendLine()
    Next
    'get a temp file name
    Dim fname As String = System.IO.Path.GetTempFileName
    'write datatable to temp file name in temp location
    Using fs As New IO.StreamWriter(fname)
      fs.Write(content.ToString)
    End Using

    Console.WriteLine("Upload CSV file to sftp site.")

    Dim fileName As String = String.Format("WS_MFI_{0}_{1}.txt",
                                          Now.ToString("yyyyMMdd"),
                                          Now.ToString("HHmm"))
    Dim sftp_uploadpath As String = SFTPUploadPath()

    Dim uploadfilename As String = IIf(sftp_uploadpath <> String.Empty, sftp_uploadpath & "/", String.Empty) _
                             & fileName

    Dim attempt As Integer = 0
    Dim done As Boolean = False
    Do
      Try
        Dim data As String = IO.File.ReadAllText(fname)
        Dim sftp_uploadsite As String = SFTPUploadSite()
        Dim sftp_uploadport As String = SFTPUploadPort()
        Dim sftp_uploadlogin As String = SFTPUploadUser()
        Dim sftp_uploadpass As String = SFTPUploadPass()

        Using client As New SftpClient(sftp_uploadsite, sftp_uploadport, sftp_uploadlogin, sftp_uploadpass)
          client.Connect()
          client.WriteAllText(uploadfilename, data)
          client.Disconnect()
        End Using

        done = True
      Catch ex As Exception
        Select Case attempt
          Case 0, 1 : System.Threading.Thread.Sleep(10000)  ' sleep 10 seconds
          Case 2 : System.Threading.Thread.Sleep(60000 * 10)  ' sleep 10 minutes
          Case 3 : System.Threading.Thread.Sleep(60000 * 20)  ' sleep 20 minutes
          Case 4 : System.Threading.Thread.Sleep(60000 * 30)  ' sleep 30 minutes
          Case Else : Throw
        End Select
        attempt += 1
      End Try
    Loop Until done

    System.IO.File.Delete(fname)
  End Sub

  Private ReadOnly Property SFTPUploadSite() As String
    Get
      Dim url As String = If(isProduction, sftp_uploadsite_prod, sftp_uploadsite_dev)
      Return url
    End Get
  End Property

  Private ReadOnly Property SFTPUploadPort() As String
    Get
      Dim port As String = If(isProduction, sftp_uploadport_prod, sftp_uploadport_dev)
      Return port
    End Get
  End Property

  Private ReadOnly Property SFTPUploadUser() As String
    Get
      Return If(isProduction, sftp_uploadlogin_prod, sftp_uploadlogin_dev)
    End Get
  End Property

  Private ReadOnly Property SFTPUploadPass() As String
    Get
      Return If(isProduction, sftp_uploadpass_prod, sftp_uploadpass_dev)
    End Get
  End Property

  Private ReadOnly Property SFTPUploadPath() As String
    Get
      Return If(isProduction, sftp_uploadpath_prod, sftp_uploadpath_dev)
    End Get
  End Property

  Public Function IsNull(ByVal data As Object, Optional ByVal nullValue As String = "") As String
    Dim result As String

    If data Is DBNull.Value Then
      result = nullValue
    ElseIf data Is Nothing Then
      result = nullValue
    Else
      result = Convert.ToString(data).Trim
    End If

    Return result
  End Function

End Module
