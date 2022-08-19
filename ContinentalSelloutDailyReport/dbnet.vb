Imports dbClassNET

Public Class dbnet
  Public Const MFIPROD As String = "mfiprod"
  Public Const MFIDEV As String = "mfidev"
  Public Const PROD As String = "R34FILES"
  Public Const DEV As String = "TESTDATA"
  Public Shared library As String = String.Empty

  Public Const dsn As String = MFIPROD

  Private Shared Function CreateConnection() As iDB2Connection
    Dim dbc As New dbClassNET.dbClassNET

    If dsn = MFIPROD Then
      dbc.setDefaultMfiprodConnection()
      library = PROD
    Else
      dbc.setDefaultMfidevConnection()
      library = DEV
    End If

    Dim cn As iDB2Connection = dbc.getConnection
    cn.Open()
    Return cn
  End Function

  Public Shared Function ContinentalSelloutData(ByVal previousDate As Integer,
                                       ByVal currentDate As Integer) As DataSet
    Dim dataset As New DataSet, dataAdapter As New iDB2DataAdapter
    Dim conn As iDB2Connection = CreateConnection()

    Try
      Dim sqlcommand As New iDB2Command(String.Format("CALL {0} . SP_CN_SELLOUTS  (@PRMPREVIOUSDATE, @PRMCURRENTDATE)", library), conn)

      sqlcommand.Parameters.Add("@PRMPREVIOUSDATE", iDB2DbType.iDB2Numeric).Value = previousDate
      sqlcommand.Parameters.Add("@PRMCURRENTDATE", iDB2DbType.iDB2Numeric).Value = currentDate
      sqlcommand.CommandTimeout = 300
      dataAdapter.SelectCommand = sqlcommand
      dataAdapter.Fill(dataset)
    Finally
      conn.Close()
    End Try

    Return dataset
  End Function
End Class
