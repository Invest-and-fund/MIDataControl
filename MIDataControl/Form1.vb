Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Form1

    Dim iSumAmount1, iSumAmount2, iSumAmount3, iSumAUM As Double

    Public Class Extract
        Property Description As String
        Property Amount1 As String
        Property Amount2 As String
        Property Amount3 As String
    End Class

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FromDate.Value = DateTime.Today.AddDays(-7)
        ToDate.Value = DateTime.Now

        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")
        TextBox1.Text = "LenderMIDataControl" & xenddate
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        LoadingMessage.Text = "Select the range of dates and presas Go"
        Me.Refresh()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        LoadingMessage.Text = "This list takes a while to load - please be patient"
        Me.Refresh()

        Dim MySQL, strConn, sHTML, bHTML, sUsers As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Dim iRegularAccounts As Integer
        Dim xAmountString As String
        Dim nAmountString As Integer
        Dim ExtractList = New List(Of Extract)
        Dim accttype As Integer

        Dim startdate As Date = FromDate.Value.ToString("yyyy/MM/dd")
        Dim enddate As Date = ToDate.Value.ToString("yyyy/MM/dd")


        Dim environ As String = "L"
        Dim connection As String = "FBConnectionString"

        Dim newExtract As New Extract

        newExtract.Description = "LENDER MI DATA CONTROL - WEEKLY UPDATE"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "Week Ending - "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "New Volumes - NUMBERS "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Lender Volumes "
        Me.Refresh()

        SetupNewVolumes(startdate, enddate, ExtractList, environ, connection, 0)
        SetupNewVolumes(startdate, enddate, ExtractList, environ, connection, 1)
        SetupNewVolumes(startdate, enddate, ExtractList, environ, connection, 2)
        SetupNewIFISA(startdate, enddate, ExtractList, environ, connection)
        SetupNewVolumesTotals(ExtractList)

        newExtract = New Extract
        newExtract.Description = "New Values "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Deposit Volumes "
        Me.Refresh()

        SetupNewDepositValues(startdate, enddate, ExtractList, environ, connection, 0)
        SetupNewDepositValues(startdate, enddate, ExtractList, environ, connection, 1)
        SetupNewDepositValues(startdate, enddate, ExtractList, environ, connection, 2)
        SetupNewDepositTotals(ExtractList)

        SetupNewWithdrawals(startdate, enddate, ExtractList, environ, connection)

        newExtract = New Extract
        newExtract.Description = "New Mandates "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Mandate Volumes "
        Me.Refresh()

        SetupNewMandates(startdate, enddate, ExtractList, environ, connection, 0)
        SetupNewMandates(startdate, enddate, ExtractList, environ, connection, 1)
        SetupNewMandates(startdate, enddate, ExtractList, environ, connection, 2)

        SetupNewMandatesTotals(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Accounts - Active "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Active Lender total "
        Me.Refresh()

        SetupTotalVolumes(ExtractList, environ, connection, 0)
        SetupTotalVolumes(ExtractList, environ, connection, 1)
        SetupTotalVolumes(ExtractList, environ, connection, 2)

        SetupTotalVolumesTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Accounts - Inctive "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Lender total "
        Me.Refresh()

        SetupInactiveVolumes(ExtractList, environ, connection, 0)
        SetupInactiveVolumes(ExtractList, environ, connection, 1)
        SetupInactiveVolumes(ExtractList, environ, connection, 2)

        SetupInactiveVolumesTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Mandates - Active "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Active Mandate total "
        Me.Refresh()

        SetupMandatesVolumes(ExtractList, environ, connection, 0)
        SetupMandatesVolumes(ExtractList, environ, connection, 1)
        SetupMandatesVolumes(ExtractList, environ, connection, 2)

        SetupMandatesVolumesTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Mandates - Inactive "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Mandate total "
        Me.Refresh()

        SetupMandatesInactive(ExtractList, environ, connection, 0)
        SetupMandatesInactive(ExtractList, environ, connection, 1)
        SetupMandatesInactive(ExtractList, environ, connection, 2)

        SetupMandatesInactiveTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Client Account Free Balances "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Client Account Free Balances"
        Me.Refresh()

        SetupFreeBalancesValues(ExtractList, environ, connection, 0)
        SetupFreeBalancesValues(ExtractList, environ, connection, 1)
        SetupFreeBalancesValues(ExtractList, environ, connection, 2)

        iSumAUM = 0

        SetupFreeBalanceTotals(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Current Loan Balances "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Current Loan Balances"
        Me.Refresh()

        SetupLoanBalancesValues(ExtractList, environ, connection, 0)
        SetupLoanBalancesValues(ExtractList, environ, connection, 1)
        SetupLoanBalancesValues(ExtractList, environ, connection, 2)

        SetupLoanBalanceTotals(ExtractList)

        SetupAUMTotals(ExtractList)

        LoadingMessage.Text = "List complete"
        Me.Refresh()



        'finally ....

        DataGridView1.DataSource = ExtractList
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80

        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")


        TextBox1.Text = "LenderMIDataControl" & xenddate
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)



    End Sub


    Private Sub SetupNewVolumes(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        MySQL = "select distinct u.userid, a.accountid
                   from users u, accounts a

                    inner join

              ( select  max (g.datecreated) as maxdatecreated, g.accountid
                    from get_user_accounts_history g
                    where  g.datecreated < @dt1
                    group by g.accountid
                    order by g.accountid   ) vt

                    on a.accountid = vt.accountid

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  = vt.accountid"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "New Trading Accounts"
            Case 1
                newExtract.Description = "New SIPP Accounts"
            Case 2
                newExtract.Description = "New ISA Accounts"
        End Select

        newExtract.Amount1 = ncount

        MySQL = "select distinct u.userid, a.accountid
                   from users u, accounts a

                    inner join

              ( select  max (g.datecreated) as maxdatecreated, g.accountid
                    from get_user_accounts_history g
                    where  g.datecreated < @dt1
                    group by g.accountid
                    order by g.accountid   ) vt

                    on a.accountid = vt.accountid

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  = vt.accountid"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        ncount = ds1.Tables(0).Rows.Count
        newExtract.Amount2 = ncount
        newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1

        Select Case accttype
            Case 0
                tbNTA1.Text = newExtract.Amount1
                tbNTA2.Text = newExtract.Amount2
                tbNTA3.Text = newExtract.Amount3
            Case 1
                tbNS1.Text = newExtract.Amount1
                tbNS2.Text = newExtract.Amount2
                tbNS3.Text = newExtract.Amount3
            Case 2
                tbNISA1.Text = newExtract.Amount1
                tbNISA2.Text = newExtract.Amount2
                tbNISA3.Text = newExtract.Amount3
        End Select

        iSumAmount1 += newExtract.Amount1
        iSumAmount2 += newExtract.Amount2
        iSumAmount3 += newExtract.Amount3


        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupNewVolumesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL NEW ACCOUNTS"
        newExtract.Amount1 = iSumAmount1
        newExtract.Amount2 = iSumAmount2
        newExtract.Amount3 = iSumAmount3

        tbTNA1.Text = iSumAmount1
        tbTNA2.Text = iSumAmount2
        tbTNA3.Text = iSumAmount3

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupNewIFISA(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct  f.accountid
                  from fin_trans f 
                  where f.transtype = 1021
                   and  f.datecreated < @dt1 "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate


        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract

        newExtract.Description = "New ISA Transfer In"


        newExtract.Amount1 = ncount

        MySQL = "select distinct  f.accountid
                  from fin_trans f 
                  where f.transtype = 1021
                   and  f.datecreated < @dt1"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate


        Adaptor.Fill(ds1)
        MyConn.Close()

        ncount = ds1.Tables(0).Rows.Count
        newExtract.Amount2 = ncount
        newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1

        tbNII1.Text = newExtract.Amount1
        tbNII2.Text = newExtract.Amount2
        tbNII3.Text = newExtract.Amount3







        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupNewDepositValues(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim dAmount1, dAmount2, dAmount3 As Double
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select sum(f.amount) as theamount 
                  from fin_trans f , accounts a
                  where f.transtype = 1100
                   and f.accountid = a.accountid
                   and a.accounttype = @at1
                   and  f.datecreated < @dt1 "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If

        Dim newExtract As New Extract


        Select Case accttype
            Case 0
                newExtract.Description = "Funds Deposited - Trading Account"
            Case 1
                newExtract.Description = "Funds Deposited - SIPP Account"
            Case 2
                newExtract.Description = "Funds Deposited - ISA Account"
        End Select


        newExtract.Amount1 = nSumm
        dAmount1 = nSumm

        MySQL = "select sum(f.amount) as theamount 
                  from fin_trans f , accounts a
                  where f.transtype = 1100
                   and f.accountid = a.accountid
                   and a.accounttype = @at1
                   and  f.datecreated < @dt1"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        nSumm = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If
        newExtract.Amount2 = nSumm
        dAmount2 = nSumm
        dAmount3 = dAmount2 - dAmount1
        newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1



        Select Case accttype
            Case 0
                tbFDT1.Text = PenceToCurrencyStringPounds(dAmount1)
                tbFDT2.Text = PenceToCurrencyStringPounds(dAmount2)
                tbFDT3.Text = PenceToCurrencyStringPounds(dAmount3)
            Case 1
                tbFDS1.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                tbFDS2.Text = PenceToCurrencyStringPounds(newExtract.Amount2)
                tbFDS3.Text = PenceToCurrencyStringPounds(newExtract.Amount3)
            Case 2
                tbFDI1.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                tbFDI2.Text = PenceToCurrencyStringPounds(newExtract.Amount2)
                tbFDI3.Text = PenceToCurrencyStringPounds(newExtract.Amount3)
        End Select

        iSumAmount1 += newExtract.Amount1
        iSumAmount2 += newExtract.Amount2
        iSumAmount3 += newExtract.Amount3

        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupNewWithdrawals(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String)
        Dim MySQL, strConn As String
        Dim dAmount1, dAmount2, dAmount3 As Double
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select sum(f.amount) as theamount 
                  from fin_trans f 
                  where f.transtype = 1102
                   and  f.datecreated < @dt1 "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate


        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If

        Dim newExtract As New Extract



        newExtract.Description = "TOTAL FUNDS WITHDRAWN IN AGGREGATE"



        newExtract.Amount1 = nSumm
        dAmount1 = nSumm

        MySQL = "select sum(f.amount) as theamount 
                  from fin_trans f 
                  where f.transtype = 1102
                   and  f.datecreated < @dt1"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate


        Adaptor.Fill(ds1)
        MyConn.Close()

        nSumm = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If
        newExtract.Amount2 = nSumm
        dAmount2 = nSumm
        dAmount3 = dAmount2 - dAmount1
        newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1




        tbTFW1.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
        tbTFW2.Text = PenceToCurrencyStringPounds(newExtract.Amount2)
        tbTFW3.Text = PenceToCurrencyStringPounds(newExtract.Amount3)


        iSumAmount1 += newExtract.Amount1
        iSumAmount2 += newExtract.Amount2
        iSumAmount3 += newExtract.Amount3

        Extractlist.Add(newExtract)
    End Sub


    Private Sub SetupNewDepositTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL FUNDS DEPOSITED"
        newExtract.Amount1 = iSumAmount1
        newExtract.Amount2 = iSumAmount2
        newExtract.Amount3 = iSumAmount3

        tbTFD1.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
        tbTFD2.Text = PenceToCurrencyStringPounds(newExtract.Amount2)
        tbTFD3.Text = PenceToCurrencyStringPounds(newExtract.Amount3)

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupNewMandates(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct   a.accountid
                from
                mandates m, accounts a
                where a.accountid = m.accountid
                and m.isactive = 0
                and a.isactive = 0
                and m.datecreated < @dt1 
                and a.accounttype = @at1"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "New Trading Mandates"
            Case 1
                newExtract.Description = "New SIPP Mandates"
            Case 2
                newExtract.Description = "New ISA Mandates"
        End Select

        newExtract.Amount1 = ncount

        MySQL = "select distinct   a.accountid
                from
                mandates m, accounts a
                where a.accountid = m.accountid
                and m.isactive = 0
                and a.isactive = 0
                and m.datecreated < @dt1 
                and a.accounttype = @at1"


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate
        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        ncount = ds1.Tables(0).Rows.Count
        newExtract.Amount2 = ncount
        newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1

        Select Case accttype
            Case 0
                tbNMT1.Text = newExtract.Amount1
                tbNMT2.Text = newExtract.Amount2
                tbNMT3.Text = newExtract.Amount3
            Case 1
                tbNMS1.Text = newExtract.Amount1
                tbNMS2.Text = newExtract.Amount2
                tbNMS3.Text = newExtract.Amount3
            Case 2
                tbNMI1.Text = newExtract.Amount1
                tbNMI2.Text = newExtract.Amount2
                tbNMI3.Text = newExtract.Amount3
        End Select

        iSumAmount1 += newExtract.Amount1
        iSumAmount2 += newExtract.Amount2
        iSumAmount3 += newExtract.Amount3


        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupNewMandatesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL NEW MANDATES"
        newExtract.Amount1 = iSumAmount1
        newExtract.Amount2 = iSumAmount2
        newExtract.Amount3 = iSumAmount3

        tbTNM1.Text = iSumAmount1
        tbTNM2.Text = iSumAmount2
        tbTNM3.Text = iSumAmount3

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupTotalVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct a.accountid
                   from users u, accounts a
                   where a.userid = u.userid and 
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  in

                (select distinct  vt.accountid
                from
                (select  a.accountid
                from users u, accounts a, lh_balances l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1

             union all
                 select  a.accountid
                from users u, accounts a, lh_balances_suspense l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1) vt      )

                group by  a.accountid
                order by  a.accountid "






        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "Total Active Trading Accounts"
            Case 1
                newExtract.Description = "Total Active SIPP Accounts"
            Case 2
                newExtract.Description = "Total Active ISA Accounts"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount



        Select Case accttype
            Case 0
                tbTAT.Text = ncount

            Case 1
                tbTAS.Text = ncount

            Case 2
                tbTAI.Text = ncount

        End Select

        iSumAmount1 += ncount



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupTotalVolumesTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL ACTIVE ACCOUNTS"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1


        tbTAA.Text = iSumAmount1


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupInactiveVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct a.accountid
                   from users u, accounts a
                   where a.userid = u.userid and 
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid not in

                (select distinct  vt.accountid
                from
                (select  a.accountid
                from users u, accounts a, lh_balances l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1

             union all
                 select  a.accountid
                from users u, accounts a, lh_balances_suspense l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1) vt      )

                group by  a.accountid
                order by  a.accountid "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "Total Inactive Trading Accounts"
            Case 1
                newExtract.Description = "Total Inactive SIPP Accounts"
            Case 2
                newExtract.Description = "Total Inactive ISA Accounts"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount



        Select Case accttype
            Case 0
                tbTIT.Text = ncount

            Case 1
                tbTIS.Text = ncount

            Case 2
                tbTII.Text = ncount

        End Select

        iSumAmount1 += ncount



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupInactiveVolumesTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL INACTIVE ACCOUNTS"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1


        tbTIA.Text = iSumAmount1


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupMandatesVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct a.accountid
                from mandates m, accounts a
                   where m.accountid = a.accountid
                  and  m.isactive = 0
                  and a.accounttype = @at1  
                group by  a.accountid
                order by  a.accountid  "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "Total Active Trading Mandates"
            Case 1
                newExtract.Description = "Total Active SIPP Mandates"
            Case 2
                newExtract.Description = "Total Active ISA Mandates"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount




        Select Case accttype
            Case 0
                tbMAT.Text = ncount

            Case 1
                tbMAS.Text = ncount

            Case 2
                tbMAI.Text = ncount

        End Select

        iSumAmount1 += ncount



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupMandatesVolumesTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL ACTIVE MANDATES"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1


        tbTAM.Text = iSumAmount1


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupMandatesInactive(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select distinct a.accountid
                from mandates m, accounts a
                   where m.accountid = a.accountid
                  and  m.isactive = 1
                  and a.accounttype = @at1  
                group by  a.accountid
                order by  a.accountid  "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "Total Inactive Trading Mandates"
            Case 1
                newExtract.Description = "Total Inactive SIPP Mandates"
            Case 2
                newExtract.Description = "Total Inactive ISA Mandates"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount



        Select Case accttype
            Case 0
                tbITM.Text = ncount

            Case 1
                tbISM.Text = ncount

            Case 2
                tbIIM.Text = ncount

        End Select

        iSumAmount1 += ncount



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupMandatesInactiveTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL INACTIVE MANDATES"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1


        tbTIM.Text = iSumAmount1


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupFreeBalancesValues(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim dAmount As Double
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet


        MySQL = "select sum(f.balance) as theamount
                from (
                 Select            fb.amount  as balance
                   From select_active_accounts a  
                  Left outer  join fin_balances fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1
              union all
                 Select            fb.amount * (-1) as balance
                   From select_active_accounts a  
                  Left outer  join fin_balances_suspense fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1) f"



        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If

        Dim newExtract As New Extract


        Select Case accttype
            Case 0
                newExtract.Description = "Total Trading Account Balances"
            Case 1
                newExtract.Description = "Total SIPP Account Balances"
            Case 2
                newExtract.Description = "Total ISA Account Balances"
        End Select



        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = nSumm



        Select Case accttype
            Case 0
                tbTAB.Text = PenceToCurrencyStringPounds(nSumm)

            Case 1
                tbSAB.Text = PenceToCurrencyStringPounds(nSumm)

            Case 2
                tbIAB.Text = PenceToCurrencyStringPounds(nSumm)

        End Select

        iSumAmount1 += nSumm


        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupFreeBalanceTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL CLIENT ACCOUNT BALANCES"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1

        tbCAB.Text = PenceToCurrencyStringPounds(newExtract.Amount1)

        iSumAUM = iSumAUM + iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupLoanBalancesValues(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String
        Dim dAmount As Double
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select sum(f.balance) as theamount
                from (
                 Select            fb.num_units  as balance
                   From select_active_accounts a  
                  Left outer  join lh_balances fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1
              union all
                 Select            fb.num_units * (-1) as balance
                   From select_active_accounts a  
                  Left outer  join lh_balances_suspense fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1) f"



        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If

        Dim newExtract As New Extract


        Select Case accttype
            Case 0
                newExtract.Description = "Trading Account Loan Balances"
            Case 1
                newExtract.Description = "SIPP Account Loan Balances"
            Case 2
                newExtract.Description = "ISA Account Loan Balances"
        End Select


        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = nSumm




        Select Case accttype
            Case 0
                tbTLB.Text = PenceToCurrencyStringPounds(nSumm)

            Case 1
                tbSLB.Text = PenceToCurrencyStringPounds(nSumm)

            Case 2
                tbILB.Text = PenceToCurrencyStringPounds(nSumm)

        End Select

        iSumAmount1 += nSumm


        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupLoanBalanceTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL LENDER LOAN BALANCES"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1

        tbLLB.Text = PenceToCurrencyStringPounds(newExtract.Amount1)

        iSumAUM = iSumAUM + iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupAUMTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL AUM (LOANS + CASH)"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAUM

        tbAUM.Text = PenceToCurrencyStringPounds(newExtract.Amount1)


        Extractlist.Add(newExtract)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim xFileName As String
        Dim xFolderPath As String

        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")



        xFileName = TextBox1.Text
        xFolderPath = TextBox2.Text
        If xFileName = "" Then
            xFileName = "LenderMIDataControl" & xenddate
        End If
        If xFolderPath = "" Then
            xFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If



        'export the files as csv using the variable name
        'Build the CSV file data as a Comma separated string.
        Dim csv As String = String.Empty

        'Add the Header row for CSV file.
        For Each column As DataGridViewColumn In DataGridView1.Columns
            csv += column.HeaderText & ","c
        Next

        'Add new line.
        csv += vbCr & vbLf

        'Adding the Rows
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                'Add the Data rows.
                csv += cell.Value.ToString().Replace(",", ";") & ","c
            Next

            'Add new line.
            csv += vbCr & vbLf
        Next

        'now add the standard entries from the config file

        Dim k As Integer = 0
        For k = 0 To 9
            Dim l As Integer = k + 1
            Dim llist As String = "l" & l
            Dim lemailstring As String = ConfigurationManager.AppSettings(llist)
            If lemailstring IsNot Nothing Then
                csv += "" & ","c
                csv += lemailstring & ","c
                csv += "" & ","c
                csv += vbCr & vbLf
            Else
                Exit For
            End If
        Next

        'Exporting to Excel

        Dim xFilePath As String
        xFilePath = xFolderPath & "\" & xFileName & ".csv"
        File.WriteAllText(xFilePath, csv)
        MessageBox.Show("CSV file written to " & xFilePath)
    End Sub



    Public Shared Function fnDBIntField(ByVal sField) As Integer
        Try
            If IsDBNull(sField) Then
                fnDBIntField = 0
            Else
                fnDBIntField = CInt(sField)
            End If
        Catch ex As Exception
            fnDBIntField = 0
        End Try
    End Function

    Public Shared Function PenceToCurrencyStringPounds(ByVal sField) As String
        Dim rVal As Double
        Dim iPence As Double

        If IsDBNull(sField) Then
            rVal = 0.0
        Else
            Try
                iPence = CDbl(sField)
                rVal = iPence / 100
            Catch ex As Exception
                rVal = 0.0
            End Try
        End If
        PenceToCurrencyStringPounds = "£" & Format(rVal, "###,###,##0.00")
    End Function
End Class
