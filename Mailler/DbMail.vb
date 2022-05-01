Imports System.Text
<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("8421A6EE-9101-4119-9262-776B6A45D28B")>
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.AutoDispatch)>
<Runtime.InteropServices.ProgId("TDS.DbMail")>
Public Class DbMail
    Implements IDbMail, IDisposable


    '	Parameters:
    '	    CN - SQL connection string
    '		smSender		sender of the e-mail
    '		smRecptLst		comma delineated string of "Rcpt To:" users
    '		smCCLst			comma delineated string of "cc:" reciepients
    '		smBCCLst		comma delineated string of "bcc:" reciepients
    '		smMailServer	name of the mail server ("titanium")
    '		smSubject		subject of the mail
    '		smMsg			actual mail message
    '       HtmlBodyFormat  True/False = HTML/TEXT, default TEXT (False)


    Public Function DbMailSendMail(CN As String, DbMailProfileName As String, SmSender As String, SmRecptLst As String, SmCCLst As String, SmBCCLst As String, SmSubject As String, SmMsg As String, SmAttach As String, Optional HtmlBodyFormat As Boolean = False) As String Implements IDbMail.DbMailSendMail
        Dim SQLCN As New SqlClient.SqlConnection(CN)
        SQLCN.Open()
        Try
            Dim CMD As New SqlClient.SqlCommand($"EXEC msdb.dbo.sp_send_dbmail @body_format = {IIf(HtmlBodyFormat, "'HTML'", "'TEXT'") }, " &
                                            $"@profile_name = '{DbMailProfileName}', " &
                                            $"@from_address='{SmSender}', " &
                                            $"@recipients = '{SmRecptLst}', " &
                                            $"@copy_recipients='{SmCCLst}', " &
                                            $"@blind_copy_recipients='{SmBCCLst}', " &
                                            $"@subject = '{SmSubject}', " &
                                            $"@body = '{SmMsg}', " &
                                            $"@file_attachments ='{SmAttach}' ", SQLCN)
            Dim Res As Integer = CMD.ExecuteNonQuery()
            SQLCN.Close()
            Return Res.ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function DbMailListProfiles(CN As String) As String Implements IDbMail.DbMailListProfiles
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_profile]", "_:_")
    End Function

    Public Function DbMailListServers(CN As String) As String Implements IDbMail.DbMailListServers
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_server]", "_:_")
    End Function

    Public Function DbMailListRerverType(CN As String) As String Implements IDbMail.DbMailListRerverType
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_servertype]", "_:_")
    End Function

    Public Function DbMailListProfileAccount(CN As String) As String Implements IDbMail.DbMailListProfileAccount
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_profileaccount]", "_:_")
    End Function

    Public Function DbMailListAccount(CN As String) As String Implements IDbMail.DbMailListAccount
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_account]", "_:_")
    End Function
    Public Function DbMailListSentEmail(CN As String) As String Implements IDbMail.DbMailListSentEmail
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_sentitems]", "_:_")
    End Function

    Public Function DbMailListUnSentEmail(CN As String) As String Implements IDbMail.DbMailListUnSentEmail
        Return GetAllFormatedDataFromSQL(CN, $"SELECT * FROM [msdb].[dbo].[sysmail_sysmail_unsentitems]", "_:_")
    End Function

    Public Function DbMailListErrors(CN As String) As String Implements IDbMail.DbMailListErrors
        Return GetAllFormatedDataFromSQL(CN, $"SELECT *,(SELECT TOP 1 [description] FROM msdb.dbo.sysmail_event_log WHERE mailitem_id = logs.mailitem_id ORDER BY log_date DESC) [description]FROM msdb.dbo.sysmail_faileditems logs", "_:_")
    End Function

    Public Function GetAllFormatedDataFromSQL(CN As String, SqlQuery As String, ColumnDivider As String) As String Implements IDbMail.GetAllFormatedDataFromSQL
        Dim SQLCN As New SqlClient.SqlConnection(CN)
        SQLCN.Open()
        Try
            Dim CMD As New SqlClient.SqlCommand(SqlQuery, SQLCN)
            Dim RDR As SqlClient.SqlDataReader = CMD.ExecuteReader
            Dim Ret1 As String = ReadAll(RDR, ColumnDivider)
            SQLCN.Close()
            Return Ret1
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Function ReadAll(RDR As SqlClient.SqlDataReader, ColumnDivider As String) As String
        Dim Sb As New StringBuilder
        For i As Integer = 0 To RDR.FieldCount - 1
            Sb.AppendFormat("{0}{1}", RDR.GetName(i), ColumnDivider)
        Next
        If RDR.HasRows Then
            While RDR.Read()
                Sb.AppendFormat($"{vbCrLf}")
                For i As Integer = 0 To RDR.FieldCount - 1
                    If Not IsDBNull(RDR.GetValue(i)) Then Sb.AppendFormat("{0}{1}", Convert.ToString(RDR.GetValue(i)), ColumnDivider) Else Sb.Append($"NULL{ColumnDivider}")
                Next
            End While
        End If
        RDR.Close()
        Return Sb.ToString
    End Function


#Region "Dispose"

    Private disposedValue As Boolean
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
