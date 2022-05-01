<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("65E58463-DB6F-421A-8DE4-A2CE106A7444")>
<Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)>
Public Interface IDbMail
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
    Function DbMailSendMail(CN As String, DbMailProfileName As String,
                  SmSender As String, SmRecptLst As String, SmCCLst As String, SmBCCLst As String, SmSubject As String, SmMsg As String, SmAttach As String, Optional HtmlBodyFormat As Boolean = False) As String
    Function GetAllFormatedDataFromSQL(CN As String, SqlQuery As String, ColumnDivider As String) As String
    Function DbMailListProfiles(CN As String) As String
    Function DbMailListServers(CN As String) As String
    Function DbMailListRerverType(CN As String) As String
    Function DbMailListAccount(CN As String) As String
    Function DbMailListProfileAccount(CN As String) As String
    Function DbMailListSentEmail(CN As String) As String
    Function DbMailListUnSentEmail(CN As String) As String
    Function DbMailListErrors(CN As String) As String


End Interface
