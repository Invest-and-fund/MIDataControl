Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Form1

    Dim iSumAmount1, iSumAmount2, iSumAmount3, iSumAUM, iTotal, iSumIUT As Double

    Public Class Extract
        Property Description As String
        Property Amount1 As String
        Property Amount2 As String
        Property Amount3 As String
    End Class

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        HorizontalScroll.Maximum = 0
        AutoScroll = False
        VerticalScroll.Visible = False
        AutoScroll = True

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

        tbNTA1.Text = ""
        tbNTA2.Text = ""
        tbNTA3.Text = ""
        tbNS1.Text = ""
        tbNS2.Text = ""
        tbNS3.Text = ""
        tbNISA1.Text = ""
        tbNISA2.Text = ""
        tbNISA3.Text = ""
        tbTNA1.Text = ""
        tbTNA2.Text = ""
        tbTNA3.Text = ""
        tbNII1.Text = ""
        tbNII2.Text = ""
        tbNII3.Text = ""
        tbFDT1.Text = ""
        tbFDT2.Text = ""
        tbFDT3.Text = ""
        tbFDS1.Text = ""
        tbFDS2.Text = ""
        tbFDS3.Text = ""
        tbFDI1.Text = ""
        tbFDI2.Text = ""
        tbFDI3.Text = ""
        tbTFW1.Text = ""
        tbTFW2.Text = ""
        tbTFW3.Text = ""
        tbTFD1.Text = ""
        tbTFD2.Text = ""
        tbTFD3.Text = ""
        tbDW.Text = ""
        tbNMT1.Text = ""
        tbNMT2.Text = ""
        tbNMT3.Text = ""
        tbNMS1.Text = ""
        tbNMS2.Text = ""
        tbNMS3.Text = ""
        tbNMI1.Text = ""
        tbNMI2.Text = ""
        tbNMI3.Text = ""
        tbTNM1.Text = ""
        tbTNM2.Text = ""
        tbTNM3.Text = ""
        tbTAT.Text = ""
        tbTASl6.Text = ""
        tbTAIl6.Text = ""
        tbTAA.Text = ""
        tbTIT.Text = ""
        tbTIS.Text = ""
        tbTII.Text = ""
        tbTIA.Text = ""
        tbMAT.Text = ""
        tbMAS.Text = ""
        tbMAI.Text = ""
        tbTAM.Text = ""
        tbITM.Text = ""
        tbISM.Text = ""
        tbIIM.Text = ""
        tbTIM.Text = ""
        tbTAB.Text = ""
        tbSAB.Text = ""
        tbIAB.Text = ""
        tbCAB.Text = ""
        tbTLB.Text = ""
        tbSLB.Text = ""
        tbILB.Text = ""
        tbAUM.Text = ""
        tbTATg6.Text = ""
        tbTATl6.Text = ""
        tbTASg6.Text = ""
        tbTASl6.Text = ""
        tbTAIg6.Text = ""
        tbTAIl6.Text = ""
        tbMATg6.Text = ""
        tbMATl6.Text = ""
        tbMASg6.Text = ""
        tbMASl6.Text = ""
        tbMAIg6.Text = ""
        tbMAIl6.Text = ""
        tbTMB.Text = ""
        tbTAS.Text = ""
        tbTAA.Text = ""




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
        newExtract.Amount3 = enddate
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

        SetupDepositWithdrawals(startdate, enddate, ExtractList, environ, connection)

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

        SetupTotalVolumes(ExtractList, environ, connection, 0, enddate)
        SetupTotalVolumes(ExtractList, environ, connection, 1, enddate)
        SetupTotalVolumes(ExtractList, environ, connection, 2, enddate)

        SetupTotalVolumesTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Accounts - Inactive "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumIUT = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Lender total "
        Me.Refresh()

        SetupInactiveVolumes(ExtractList, environ, connection, 0, enddate)
        SetupInactiveVolumes(ExtractList, environ, connection, 1, enddate)
        SetupInactiveVolumes(ExtractList, environ, connection, 2, enddate)

        SetupInactiveVolumesTotal(ExtractList)
        newExtract = New Extract
        newExtract.Description = "Lender Accounts - Unfunded "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0


        LoadingMessage.Text = "This list takes a while to load - calculating Unfunded Lender total "
        Me.Refresh()

        SetupUnfundedVolumes(ExtractList, environ, connection, 0, enddate)
        SetupUnfundedVolumes(ExtractList, environ, connection, 1, enddate)
        SetupUnfundedVolumes(ExtractList, environ, connection, 2, enddate)

        SetupUnfundedVolumesTotal(ExtractList)
        SetupInactiveUnfundedTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Lender Mandates - Active "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Active Mandate total "
        Me.Refresh()

        SetupMandatesVolumes(ExtractList, environ, connection, 0, enddate)
        SetupMandatesVolumes(ExtractList, environ, connection, 1, enddate)
        SetupMandatesVolumes(ExtractList, environ, connection, 2, enddate)

        SetupMandatesVolumesTotal(ExtractList)

        SetupActiveMandatesBalance(ExtractList, environ, connection, enddate)

        newExtract = New Extract
        newExtract.Description = "Lender Mandates - Inactive "
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Mandate total "
        Me.Refresh()

        SetupMandatesInactive(ExtractList, environ, connection, 0, enddate)
        SetupMandatesInactive(ExtractList, environ, connection, 1, enddate)
        SetupMandatesInactive(ExtractList, environ, connection, 2, enddate)

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

        SetupFreeBalancesValues(ExtractList, environ, connection, 0, enddate)
        SetupFreeBalancesValues(ExtractList, environ, connection, 1, enddate)
        SetupFreeBalancesValues(ExtractList, environ, connection, 2, enddate)

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

        SetupLoanBalancesValues(ExtractList, environ, connection, 0, enddate)
        SetupLoanBalancesValues(ExtractList, environ, connection, 1, enddate)
        SetupLoanBalancesValues(ExtractList, environ, connection, 2, enddate)

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
                  from fin_trans f , accounts a, users u
                  where f.transtype = 1100
                   and f.accountid = a.accountid
                   and a.accounttype = @at1
                   and  f.datecreated <= @dt1
                   and f.isactive = 0                  
                   and f.accountid   >= 2  
                   and f.accountid   <> 20 
                   and f.accountid = a.accountid
                   and a.userid = u.userid                   
                   and u.usertype <> 1"


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
                  from fin_trans f , accounts a, users u
                  where f.transtype = 1100
                   and f.accountid = a.accountid
                   and a.accounttype = @at1
                   and  f.datecreated <= @dt1
                   and f.isactive = 0                  
                   and f.accountid   >= 2  
                   and f.accountid   <> 20 
                   and f.accountid = a.accountid
                   and a.userid = u.userid                   
                   and u.usertype <> 1"


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
                  from fin_trans f, accounts a, users u 
                  where f.transtype = 1102
                   and  f.datecreated <= @dt1
                   and f.isactive = 0                  
                   and f.accountid   >= 2  
                   and f.accountid   <> 20 
                   and f.accountid = a.accountid
                   and a.userid = u.userid                   
                   and u.usertype <> 1"


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
                  from fin_trans f, accounts a, users u 
                  where f.transtype = 1102
                   and  f.datecreated <= @dt1
                   and f.isactive = 0                  
                   and f.accountid   >= 2  
                   and f.accountid   <> 20 
                   and f.accountid = a.accountid
                   and a.userid = u.userid                   
                   and u.usertype <> 1"


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

    Private Sub SetupDepositWithdrawals(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String)
        Dim MySQL, strConn As String
        Dim dAmount1, dAmount2, dAmount3 As Double
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        MySQL = "select sum(f.amount) as theamount 
                  from fin_trans f , accounts a, users u
                  where f.transtype = 1102
                   and f.accountid = a.accountid
                   and f.isactive = 0
                   and  f.datecreated > @dt1
                   and f.accountid   >= 2  
                   and f.accountid   <> 20 
                   and a.userid = u.userid                   
                   and u.usertype <> 1
                   and f.fin_transid in
              (  select g.fin_transid
                  from fin_trans f , fin_trans g
                  where f.transtype = 1100
                  and g.transtype = 1102
                  and f.accountid = g.accountid
                  and f.amount = g.amount
                  and f.datecreated > dateadd(day, -7, g.datecreated )
                  and g.datecreated > @dt1
                  and g.datecreated < @dt2  ) "




        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = startdate
        Adaptor.SelectCommand.Parameters.Add("@dt2", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate


        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                If Not IsDBNull(ds1.Tables(0).Rows(0).Item(“theamount”)) Then
                    nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
                Else
                    nSumm = 0
                End If

            End With
        End If

        Dim newExtract As New Extract



        newExtract.Description = "Funds withdrawn within one week of deposit"


        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = nSumm






        tbDW.Text = PenceToCurrencyStringPounds(newExtract.Amount3)




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

    Private Sub SetupTotalVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)

        iTotal = 0
        Dim eLoop As Integer = 0
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, enddate, eLoop, iTotal)

        eLoop = 1
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, enddate, eLoop, iTotal)

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
        newExtract.Amount3 = iTotal



        Select Case accttype
            Case 0

                tbTAT.Text = iTotal

            Case 1

                tbTAS.Text = iTotal

            Case 2

                tbTAI.Text = iTotal

        End Select

        'iSumAmount1 += nTotal



        Extractlist.Add(newExtract)



    End Sub
    Private Sub SetupTotalVolumes6mth(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date, eLoop As Integer, nTotal As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Dim eSQL As String

        If eLoop = 0 Then
            eSQL = " and not exists "
        Else

            eSQL = " and exists "
        End If
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

                and a.accountid in

                  (select distinct a.accountid
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
                     and a.accountid  = vt.accountid  " & eSQL &
             "  (select distinct accountid from lh_bals t
                    where t.num_units > 0
                      and t.accountid = a.accountid
                      and t.datecreated > dateadd(month,  -3, @dt1)

                      union all
               select distinct accountid from lh_bals_suspense t
                    where t.num_units > 0
                      and t.accountid = a.accountid
                      and t.datecreated > dateadd(month,  -3, @dt1)
                         ))

                group by  a.accountid
                order by  a.accountid "






        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                If eLoop = 0 Then
                    newExtract.Description = "Active Trading Accounts - Live"
                Else
                    newExtract.Description = "Active Trading Accounts - Currently Lending"
                End If

            Case 1
                If eLoop = 0 Then
                    newExtract.Description = "Active SIPP Accounts - Live"
                Else
                    newExtract.Description = "Active SIPP Accounts - Currently Lending"
                End If

            Case 2
                If eLoop = 0 Then
                    newExtract.Description = "Active ISA Accounts - Live"
                Else
                    newExtract.Description = "Active ISA Accounts - Currently Lending"
                End If

        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount



        Select Case accttype
            Case 0
                If eLoop = 0 Then
                    tbTATl6.Text = ncount
                Else
                    tbTATg6.Text = ncount
                End If


            Case 1
                If eLoop = 0 Then
                    tbTASl6.Text = ncount
                Else
                    tbTASg6.Text = ncount
                End If


            Case 2
                If eLoop = 0 Then
                    tbTAIl6.Text = ncount
                Else
                    tbTAIg6.Text = ncount
                End If


        End Select

        iSumAmount1 += ncount

        iTotal += ncount

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

    Private Sub SetupInactiveVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
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

               and a.accountid in

                  (select distinct a.accountid
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
                     and a.accountid  = vt.accountid)

                group by  a.accountid
                order by  a.accountid "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

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
        iSumIUT += iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupUnfundedVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
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
                from  lh_bals l
                  where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0
                  and l.datecreated < @dt1
                  and a.accounttype = @at1  ) vt      )



                group by  a.accountid
                order by  a.accountid "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "Total Unfunded Trading Accounts"
            Case 1
                newExtract.Description = "Total Unfunded SIPP Accounts"
            Case 2
                newExtract.Description = "Total Unfunded ISA Accounts"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount



        Select Case accttype
            Case 0
                tbTUT.Text = ncount

            Case 1
                tbTUS.Text = ncount

            Case 2
                tbTUI.Text = ncount

        End Select

        iSumAmount1 += ncount



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupUnfundedVolumesTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL UNFUNDED ACCOUNTS"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAmount1


        tbTUA.Text = iSumAmount1
        iSumIUT += iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupInactiveUnfundedTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL INACTIVE AND UNFUNDED ACCOUNTS"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumIUT


        tbIUA.Text = iSumIUT


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupMandatesVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
        iTotal = 0
        Dim eLoop As Integer = 0
        SetupMandatesVolumes6mth(Extractlist, environ, connection, accttype, enddate, eLoop, iTotal)

        eLoop = 1
        SetupMandatesVolumes6mth(Extractlist, environ, connection, accttype, enddate, eLoop, iTotal)

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
        newExtract.Amount3 = iTotal




        Select Case accttype
            Case 0
                tbMAT.Text = iTotal

            Case 1
                tbMAS.Text = iTotal

            Case 2
                tbMAI.Text = iTotal

        End Select

        'iSumAmount1 += nTotal



        Extractlist.Add(newExtract)
    End Sub

    Private Sub SetupMandatesVolumes6mth(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date, eLoop As Integer, nTotal As Integer)
        Dim MySQL, strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Dim eSQL As String
        If eLoop = 0 Then
            eSQL = " and not exists "
        Else

            eSQL = " and exists "
        End If
        MySQL = "select distinct a.accountid
                from mandates m, accounts a
                   where m.accountid = a.accountid
                  and  m.isactive = 0
                  and a.accounttype = @at1 

                and a.accountid in

                  (select distinct a.accountid
                   from users u, accounts a

                    inner join

              ( select  max (g.datecreated) as maxdatecreated, g.accountid
                    from get_user_accounts_history g
                    where  g.datecreated < @dt1
                    group by g.accountid
                    order by g.accountid   ) vt

                    on a.accountid = vt.accountid  " & eSQL &
           "  (
                     select distinct accountid from lh_bals t
                    where t.num_units > 0
                      and t.accountid = a.accountid
                      and t.datecreated > dateadd(month,  -3, @dt1)

                      union all
               select distinct accountid from lh_bals_suspense t
                    where t.num_units > 0
                      and t.accountid = a.accountid
                      and t.datecreated > dateadd(month,  -3, @dt1)

                     )        

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  = vt.accountid)

                group by  a.accountid
                order by  a.accountid  "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim ncount As Integer = ds1.Tables(0).Rows.Count


        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                If eLoop = 0 Then
                    newExtract.Description = "Active Trading Mandates - Active"
                Else
                    newExtract.Description = "Active Trading Mandates - Currently Lending"
                End If

            Case 1
                If eLoop = 0 Then
                    newExtract.Description = "Active SIPP Mandates - Active"
                Else
                    newExtract.Description = "Active SIPP Mandates - Currently Lending"
                End If

            Case 2
                If eLoop = 0 Then
                    newExtract.Description = "Active ISA Mandates - Active"
                Else
                    newExtract.Description = "Active ISA Mandates - Currently Lending"
                End If

        End Select


        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = ncount

        Select Case accttype
            Case 0
                If eLoop = 0 Then
                    tbMATl6.Text = ncount
                Else
                    tbMATg6.Text = ncount
                End If


            Case 1
                If eLoop = 0 Then
                    tbMASl6.Text = ncount
                Else
                    tbMASg6.Text = ncount
                End If


            Case 2
                If eLoop = 0 Then
                    tbMAIl6.Text = ncount
                Else
                    tbMAIg6.Text = ncount
                End If
        End Select




        iSumAmount1 += ncount

        iTotal += ncount

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

    Private Sub SetupActiveMandatesBalance(Extractlist As List(Of Extract), environ As String, connection As String, enddate As Date)
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
        inner join
          ( select distinct a.accountid
                from mandates m, accounts a
                where m.accountid = a.accountid
                 and m.isactive = 0
                   ) mt on a.accountid = mt.accountid
                 Left outer  join
                 (select vt.accountid,  max_fin_balid, t.amount
        from 
        ( 
        select accountid, max(fin_balid) as max_fin_balid
        from fin_bals s

        where s.datecreated < @dt1
        group by accountid
        ) vt
        inner join fin_bals t on t.fin_balid = vt.max_fin_balid
        where t.amount > 0 ) fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
              union all
                 Select            fb.amount as balance
                   From select_active_accounts a
        inner join
          ( select distinct a.accountid
                from mandates m, accounts a
                where m.accountid = a.accountid
                 and m.isactive = 0
                   ) mt on a.accountid = mt.accountid
                  Left outer  join
            (select vt.accountid,  max_fin_bals_suspenseid, t.amount
        from 
        ( 
        select accountid, max(fin_bals_suspenseid) as max_fin_bals_suspenseid
        from fin_bals_suspense s

        where s.datecreated < @dt1
        group by accountid
        ) vt
        inner join fin_bals_suspense t on t.fin_bals_suspenseid = vt.max_fin_bals_suspenseid
        where t.amount > 0 )   fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid    ) f"



        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)


        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

        Adaptor.Fill(ds1)
        MyConn.Close()

        Dim nSumm As Double = 0
        If ds1.Tables(0).Rows.Count > 0 Then
            With ds1.Tables(0).Rows(0)
                nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
            End With
        End If
        Dim newExtract As New Extract

        newExtract.Description = "TOTAL ACTIVE MANDATES BALANCES"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = nSumm


        tbTMB.Text = PenceToCurrencyStringPounds(nSumm)


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupMandatesInactive(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
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

                and a.accountid in

                  (select distinct a.accountid
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
                     and a.accountid  = vt.accountid)

                group by  a.accountid
                order by  a.accountid  "


        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

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

    Private Sub SetupFreeBalancesValues(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
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
                  Left outer  join

                 (select vt.accountid,  max_fin_balid, t.amount
        from 
        ( 
        select accountid, max(fin_balid) as max_fin_balid
        from fin_bals s

        where s.datecreated < @dt1
        group by accountid
        ) vt
        inner join fin_bals t on t.fin_balid = vt.max_fin_balid
        where t.amount > 0 ) fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1
              union all
                 Select            fb.amount as balance
                   From select_active_accounts a  
                  Left outer  join

            (select vt.accountid,  max_fin_bals_suspenseid, t.amount
        from 
        ( 
        select accountid, max(fin_bals_suspenseid) as max_fin_bals_suspenseid
        from fin_bals_suspense s

        where s.datecreated < @dt1
        group by accountid
        ) vt
        inner join fin_bals_suspense t on t.fin_bals_suspenseid = vt.max_fin_bals_suspenseid
        where t.amount > 0 )   fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1) f"



        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

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

        tbCAB.Text = PenceToCurrencyStringPounds(iSumAmount1)

        iSumAUM = iSumAUM + iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupLoanBalancesValues(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
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
                  Left outer  join
                  ( select vt.accountid, lh_id, max_lh_bals_id, t.num_units
                 from 
                  ( 
                     select accountid, max(lh_bals_id) as max_lh_bals_id
                     from lh_bals s
                      where lh_id > 0
                       and s.datecreated < @dt1
                       group by accountid, lh_id
                    ) vt
                inner join lh_bals t on t.lh_bals_id = vt.max_lh_bals_id
                 where t.num_units > 0 ) 

                fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1
              union all
                 Select            fb.num_units  as balance
                   From select_active_accounts a  
                  Left outer  join 
            (select vt.accountid, lh_id, max_lh_bals_sus_id, t.num_units
                 from 
                  ( 
                     select accountid, max(lh_bals_suspense_id) as max_lh_bals_sus_id
                     from lh_bals_suspense s
                      where lh_id > 0
                       and s.datecreated < @dt1
                       group by accountid, lh_id
                    ) vt
                inner join lh_bals_suspense t on t.lh_bals_suspense_id = vt.max_lh_bals_sus_id
                 where t.num_units > 0 )   fb on fb.accountid = a.accountid
                 where fb.accountid = a.accountid
                   and a.accounttype = @at1) f"



        strConn = ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        ds1 = New DataSet
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@at1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = accttype
        Adaptor.SelectCommand.Parameters.Add("@dt1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = enddate

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

        tbLLB.Text = PenceToCurrencyStringPounds(iSumAmount1)

        iSumAUM = iSumAUM + iSumAmount1

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupAUMTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL AUM (LOANS + CASH)"
        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iSumAUM

        tbAUM.Text = PenceToCurrencyStringPounds(iSumAUM)


        Extractlist.Add(newExtract)



    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles tbTASg6.TextChanged

    End Sub

    Private Sub Label61_Click(sender As Object, e As EventArgs) Handles Label61.Click

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click

    End Sub

    Private Sub Label68_Click(sender As Object, e As EventArgs) Handles Label68.Click

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
