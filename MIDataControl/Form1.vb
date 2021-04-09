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

        Dim s, strConn As String
        s = ""
        Dim [Start], [End] As Integer
        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString.ToUpper
        [Start] = strConn.IndexOf("DATABASE=", 0) + "DATABASE=".Length
        [End] = strConn.IndexOf(";", Start)
        s = strConn.Substring(Start, [End] - Start)
        If s = "MAIN" Then
            s = "Live SQL Server"
        Else
            s = s + " Server"
        End If
        Me.Text = Me.Text + " - " + s

        FromDate.Value = DateTime.Today.AddDays(-7)
        ToDate.Value = DateTime.Now

        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")
        TextBox1.Text = "LenderMIDataControl" & xenddate
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        LoadingMessage.Text = "Select the range of dates and press Go"
        Me.Refresh()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        LoadingMessage.Text = "This list takes a while to load - please be patient"


        tbNTA3.Text = ""

        tbNS3.Text = ""

        tbNISA3.Text = ""

        tbTNA3.Text = ""

        tbNMT3.Text = ""

        tbNMS3.Text = ""

        tbNMI3.Text = ""

        tbTNM3.Text = ""
        tbTAT.Text = ""
        tbTASl3.Text = ""
        tbTAIl3.Text = ""
        tbTAA.Text = ""
        tbTIT.Text = ""
        tbTIS.Text = ""
        tbTII.Text = ""
        tbTIA.Text = ""
        'tbMAT.Text = ""
        'tbMAS.Text = ""
        'tbMAI.Text = ""
        tbTAM.Text = ""

        tbTAB.Text = ""
        tbSAB.Text = ""
        tbIAB.Text = ""
        tbCAB.Text = ""
        tbTLB.Text = ""
        tbSLB.Text = ""
        tbILB.Text = ""
        tbAUM.Text = ""
        tbTATg12.Text = ""
        tbTATl3.Text = ""
        tbTASl6.Text = ""
        tbTASl3.Text = ""
        tbTAIl6.Text = ""
        tbTAIl3.Text = ""
        tbMATg6.Text = ""
        tbMATl6.Text = ""
        tbMASg6.Text = ""

        tbMAIg6.Text = ""

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

        SetupNewVolumesTotals(ExtractList)



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

        '''this is where the active totals go

        SetupTotalVolumes(ExtractList, environ, connection, 0)
        SetupTotalVolumes(ExtractList, environ, connection, 1)
        SetupTotalVolumes(ExtractList, environ, connection, 2)

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

        SetupMandatesVolumes(ExtractList, environ, connection, 9, enddate)
        SetupMandatesVolumes(ExtractList, environ, connection, 0, enddate)
        SetupMandatesVolumes(ExtractList, environ, connection, 1, enddate)
        SetupMandatesVolumes(ExtractList, environ, connection, 2, enddate)

        SetupMandatesVolumesTotal(ExtractList)

        SetupActiveMandatesBalance(ExtractList, environ, connection, enddate)



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

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct u.userid, a.accountid
                   from users u, accounts a

                    inner join

              ( select  max (g.datecreated) as maxdatecreated, g.accountid
                    from get_user_accounts_history g
                    where  g.datecreated < @dt1
                    group by g.accountid   ) vt
               

                    on a.accountid = vt.accountid

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@at1", accttype))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

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

                adapter = New SqlDataAdapter()
                MySQL = "select distinct u.userid, a.accountid
                   from users u, accounts a

                    inner join

              ( select  max (g.datecreated) as maxdatecreated, g.accountid
                    from get_user_accounts_history g
                    where  g.datecreated < @dt1
                    group by g.accountid   ) vt
  

                    on a.accountid = vt.accountid

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  = vt.accountid"


                cmd = New SqlCommand(MySQL, con)

                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()

                ncount = ds1.Tables(0).Rows.Count
                newExtract.Amount2 = ncount
                newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1

                Select Case accttype
                    Case 0

                        tbNTA3.Text = newExtract.Amount3
                    Case 1

                        tbNS3.Text = newExtract.Amount3
                    Case 2

                        tbNISA3.Text = newExtract.Amount3
                End Select


                iSumAmount3 += newExtract.Amount3


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupNewVolumesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL NEW ACCOUNTS"

        newExtract.Amount3 = iSumAmount3


        tbTNA3.Text = iSumAmount3

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupNewIFISA(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct  f.accountid
                  from fin_trans f 
                  where f.transtype = 1021
                   and  f.datecreated < @dt1 "

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract

                newExtract.Description = "New ISA Transfer In"


                newExtract.Amount1 = ncount

                adapter = New SqlDataAdapter()
                MySQL = "select distinct  f.accountid
                  from fin_trans f 
                  where f.transtype = 1021
                   and  f.datecreated < @dt1"

                cmd = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", enddate))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()

                ncount = ds1.Tables(0).Rows.Count
                newExtract.Amount2 = ncount
                newExtract.Amount3 = newExtract.Amount2 - newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub




    Private Sub SetupNewMandates(startdate As Date, enddate As Date, Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet



        Dim newExtract As New Extract
        Select Case accttype
            Case 0
                newExtract.Description = "New Standard Mandates"
            Case 1
                newExtract.Description = "New SIPP Mandates"
            Case 2
                newExtract.Description = "New ISA Mandates"
        End Select


        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "Select distinct   a.accountid
                from
                mandates m, accounts a
                where a.accountid = m.accountid
                And m.isactive = 0
                And a.isactive = 0
                And m.datecreated < @dt1 
                And a.accounttype = @at1
                And m.accountid Not in
                (select distinct   n.accountid
                from
                mandates n
                where m.datecreated < @dt2)"

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt2", startdate))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()

                Dim ncount As Integer = ds1.Tables(0).Rows.Count
                newExtract.Amount3 = ncount


                Select Case accttype
                    Case 0

                        tbNMT3.Text = newExtract.Amount3
                    Case 1

                        tbNMS3.Text = newExtract.Amount3
                    Case 2

                        tbNMI3.Text = newExtract.Amount3
                End Select


                iSumAmount3 += newExtract.Amount3


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub

    Private Sub SetupNewMandatesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL NEW MANDATES"

        newExtract.Amount3 = iSumAmount3


        tbTNM3.Text = iSumAmount3

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupTotalVolumes(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer)

        iTotal = 0
        Dim enddate As Date = Date.Now()
        Dim startdate As Date = enddate.AddMonths(-3)
        Dim eloop As Integer = 0
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, startdate, enddate, iTotal, eloop)

        enddate = startdate
        startdate = enddate.AddMonths(-3)
        eloop = 1
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, startdate, enddate, iTotal, eloop)

        enddate = startdate
        startdate = enddate.AddMonths(-6)
        eloop = 2
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, startdate, enddate, iTotal, eloop)

        enddate = startdate
        startdate = enddate.AddMonths(-12)
        eloop = 3
        SetupTotalVolumes6mth(Extractlist, environ, connection, accttype, startdate, enddate, iTotal, eloop)

        Dim newExtract As New Extract
        Select Case accttype
            Case 0

                newExtract.Description = "Total Active Standard Accounts"
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
    Private Sub SetupTotalVolumes6mth(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, startdate As Date, enddate As Date, nTotal As Integer, eloop As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Dim eSQL As String
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct a.accountid
                   from users u, accounts a
                   where a.userid = u.userid and 
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  in

                (select  distinct  vt.accountid
                from
                (select  a.accountid
                from users u, accounts a, lh_bals l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1
                  and l.datecreated > @dt2
                  and l.datecreated <= @dt1   ) vt  )

                  and a.accountid not in
               (select  distinct  wt.accountid
                from
                (select  a.accountid
                from users u, accounts a, lh_bals l
                   where u.userid = a.userid
                  and  u.activated = 5
                  and a.activated_bank = 5
                  and (u.activated_cert = 5 or u.veteran_1914 = 0)  
                  and u.isactive = 0
                  and a.isactive = 0
                  and a.accountid = l.accountid
                  and l.num_units > 0  
                  and a.accounttype = @at1

                  and l.datecreated > @dt1   ) wt  )

                group by  a.accountid
                order by  a.accountid "

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt1", enddate))
                    .Add(New SqlParameter("@dt2", startdate))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                Select Case accttype
                    Case 0
                        Select Case eloop
                            Case 0
                                newExtract.Description = "Active Standard Accounts < 3 Months"
                            Case 1
                                newExtract.Description = "Active Standard Accounts < 6 Months"
                            Case 2
                                newExtract.Description = "Active Standard Accounts < 12 Months"
                            Case 3
                                newExtract.Description = "Active Standard Accounts > 12 Months"
                        End Select

                    Case 1
                        Select Case eloop
                            Case 0
                                newExtract.Description = "Active SIPP Accounts < 3 Months"
                            Case 1
                                newExtract.Description = "Active SIPP Accounts < 6 Months"
                            Case 2
                                newExtract.Description = "Active SIPP Accounts < 12 Months"
                            Case 3
                                newExtract.Description = "Active SIPP Accounts > 12 Months"
                        End Select


                    Case 2
                        Select Case eloop
                            Case 0
                                newExtract.Description = "Active ISA Accounts < 3 Months"
                            Case 1
                                newExtract.Description = "Active ISA Accounts < 6 Months"
                            Case 2
                                newExtract.Description = "Active ISA Accounts < 12 Months"
                            Case 3
                                newExtract.Description = "Active ISA Accounts > 12 Months"
                        End Select


                End Select

                newExtract.Amount1 = ""
                newExtract.Amount2 = ""
                newExtract.Amount3 = ncount



                Select Case accttype
                    Case 0
                        Select Case eloop
                            Case 0
                                tbTATl3.Text = ncount
                            Case 1
                                tbTATl6.Text = ncount
                            Case 2
                                tbTATl12.Text = ncount
                            Case 3
                                tbTATg12.Text = ncount
                        End Select



                    Case 1
                        Select Case eloop
                            Case 0
                                tbTASl3.Text = ncount
                            Case 1
                                tbTASl6.Text = ncount
                            Case 2
                                tbTASl12.Text = ncount
                            Case 3
                                tbTASg12.Text = ncount
                        End Select



                    Case 2
                        Select Case eloop
                            Case 0
                                tbTAIl3.Text = ncount
                            Case 1
                                tbTAIl6.Text = ncount
                            Case 2
                                tbTAIl12.Text = ncount
                            Case 3
                                tbTAIg12.Text = ncount
                        End Select



                End Select

                iSumAmount1 += ncount

                iTotal += ncount

                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
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

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
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
                    group by g.accountid   ) vt
      

                    on a.accountid = vt.accountid

               where a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0
                     and a.accounttype = @at1
                     and a.accountid  = vt.accountid)

                group by  a.accountid "

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt1", enddate))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()




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
            Catch ex As Exception
            Finally

            End Try
        End Using
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

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
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

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt1", enddate))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()




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
            Catch ex As Exception
            Finally

            End Try
        End Using
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


        SetupMandatesVolumesBkdn(Extractlist, environ, connection, accttype, enddate, iTotal)

        Dim newExtract As New Extract



        Select Case accttype
            Case 0
                newExtract.Description = "Total Active Trading Mandates"
            Case 1
                newExtract.Description = "Total Active SIPP Mandates"
            Case 2
                newExtract.Description = "Total Active ISA Mandates"
            Case 9
                newExtract.Description = "Active Mandates"
        End Select

        newExtract.Amount1 = ""
        newExtract.Amount2 = ""
        newExtract.Amount3 = iTotal

        Extractlist.Add(newExtract)
    End Sub



    Private Sub SetupMandatesVolumesBkdn(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date, nTotal As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Dim eSQL1, eSQL2 As String
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                If accttype = 9 Then
                    eSQL1 = " not "
                    eSQL2 = " "
                Else
                    eSQL2 = " and a.accounttype = @at1 "
                    eSQL1 = " "
                End If
                MySQL = "select distinct a.accountid
                from mandates m, accounts a , users u
                   where m.accountid = a.accountid
                  and  m.isactive = 0 " & eSQL2 &
                         " and a.userid = u.userid and
                     u.activated = 5 and a.activated_bank = 5 
                     and u.isactive = 0  And  u.usertype = 0

                and a.accountid " & eSQL1 & "  in



              (
                     select distinct accountid from lh_balances t
                    where t.num_units > 0
                      and t.accountid = a.accountid


                      union all
               select distinct accountid from lh_balances_suspense t
                    where t.num_units > 0
                      and t.accountid = a.accountid


                      union all
                      select distinct accountid from fin_balances t
                      where t.amount > 0
                      and t.accountid = a.accountid

                      union all
                      select distinct accountid from fin_balances_suspense t
                      where t.amount > 0
                      and t.accountid = a.accountid

                     )


                group by  a.accountid
                order by  a.accountid  "

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()


                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                Select Case accttype
                    Case 0

                        newExtract.Description = "Active Trading Mandates - Currently Lending"


                    Case 1

                        newExtract.Description = "Active SIPP Mandates - Currently Lending"


                    Case 2

                        newExtract.Description = "Active ISA Mandates - Currently Lending"

                    Case 9

                        newExtract.Description = "Active Mandates -Live No Funds"

                End Select


                newExtract.Amount1 = ""
                newExtract.Amount2 = ""
                newExtract.Amount3 = ncount

                Select Case accttype
                    Case 0

                        tbMATg6.Text = ncount



                    Case 1

                        tbMASg6.Text = ncount



                    Case 2

                        tbMAIg6.Text = ncount

                    Case 9
                        tbMATl6.Text = ncount

                End Select




                iSumAmount1 += ncount

                iTotal += ncount

                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
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

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
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

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", enddate))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()

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

            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub


    Private Sub SetupFreeBalancesValues(Extractlist As List(Of Extract), environ As String, connection As String, accttype As Integer, enddate As Date)
        Dim MySQL, strConn As String
        Dim dAmount As Double

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
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

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt1", enddate))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()



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
            Catch ex As Exception
            Finally

            End Try
        End Using
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

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
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

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt1", enddate))

                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                con.Close()



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
            Catch ex As Exception
            Finally

            End Try
        End Using
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

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles tbTASl6.TextChanged

    End Sub

    Private Sub Label61_Click(sender As Object, e As EventArgs) Handles Label61.Click

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click

    End Sub

    Private Sub Label68_Click(sender As Object, e As EventArgs)

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
