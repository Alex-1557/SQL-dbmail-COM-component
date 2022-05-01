Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1
    Dim CN = "Data Source=162.251.147.189; Integrated Security=SSPI"

    Private TestContextInstance As TestContext
    Public Property TestContext As TestContext
        Get
            Return TestContextInstance
        End Get
        Set(ByVal value As TestContext)
            TestContextInstance = value
        End Set
    End Property

    <TestMethod()>
    Public Sub ListProfilesTest()
        Dim X As New TdsMail.DbMail
        TestContext.WriteLine(X.DbMailListProfiles(CN))
    End Sub
    <TestMethod()>
    Public Sub ListServersTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListServers(CN))
    End Sub
    <TestMethod()>
    Public Sub ListRerverTypeTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListRerverType(CN))
    End Sub
    <TestMethod()>
    Public Sub ListAccountTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListAccount(CN))
    End Sub
    <TestMethod()>
    Public Sub ListProfileAccountTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListProfileAccount(CN))
    End Sub
    <TestMethod()>
    Public Sub ListSentEmailTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListSentEmail(CN))
    End Sub
    <TestMethod()>
    Public Sub ListUnSentEmailTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListUnSentEmail(CN))
    End Sub
    <TestMethod()>
    Public Sub ListErrorsTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailListErrors(CN))
    End Sub

    <TestMethod()>
    Public Sub GetAllFormatedDataFromSQLTest()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.GetAllFormatedDataFromSQL(CN, "Select * from [TDSv5].[dbo].[Security_Users]", " : "))
    End Sub
    <TestMethod()>
    Public Sub SendTest1()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailSendMail(CN, "365", "", "H...B...@gmail.com", "", "", "newline3", "T2" & vbCrLf & "T2", ""))
    End Sub

    <TestMethod()>
    Public Sub SendTest2()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailSendMail(CN, "365", "", "H...B...@gmail.com", "", "", "newline4", "T2" & vbCrLf & "T2", "", False))
    End Sub

    <TestMethod()>
    Public Sub SendTest3()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailSendMail(CN, "365", "", "H...B...@gmail.com", "", "", "T3", "T3", "C:\www\tdsv5\Images\BeOS_Help.gif"))
    End Sub

    <TestMethod()>
    Public Sub SendTest4()
        Dim X As New TDSMail.DbMail
        TestContext.WriteLine(X.DbMailSendMail(CN, "365", "", "H...B...@gmail.com", "", "", "T2", "T2", ""))
    End Sub

End Class