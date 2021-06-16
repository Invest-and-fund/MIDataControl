Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Form2
    Dim iSumAmount1, iSumAmount2, iSumAmount3, iSumAmount4, iSumAUM, iTotal, iSumIUT As Double
    Dim iSSThreshold As Integer

    Public Class Extract
        Property Description As String
        Property Amount1 As String

    End Class

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'AutoScroll = False
        'HorizontalScroll.Enabled = False
        ''   HorizontalScroll.Maximum = 0
        'VerticalScroll.Enabled = False
        'VerticalScroll.Visible = True
        'AutoScroll = True



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
        Me.Text = "MI Reporting - " + s

        FromDate.Value = DateTime.Today.AddDays(-7)
        ToDate.Value = DateTime.Now

        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")
        TextBox1.Text = "LenderMIDataControl" & xenddate
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        LoadingMessage.Text = "Select the range of dates and press Go"

        ssThreshold.Text = PenceToCurrencyStringPounds(2500000)
        Me.Refresh()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        LoadingMessage.Text = "This list takes a while to load - please be patient"

        iSSThreshold = CurrencyStringPoundsToPence(ssThreshold.Text)

        Me.Refresh()

        Dim MySQL, strConn, sHTML, bHTML, sUsers As String

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

        ExtractList.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "Week Ending - "

        newExtract.Amount1 = enddate
        ExtractList.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "
        newExtract.Amount1 = ""

        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount3 = 0
        iSumAmount2 = 0


        LoadingMessage.Text = "This list takes a while to load - calculating Mandatory Mandate Volumes "
        Me.Refresh()

        SetupMandatoryMandates(startdate, enddate, ExtractList, 0)
        SetupMandatoryMandates(startdate, enddate, ExtractList, 1)
        SetupMandatoryMandates(startdate, enddate, ExtractList, 2)

        SetupMandatoryMandatesTotals(ExtractList)



        newExtract = New Extract
        newExtract.Description = "Self-Select "
        newExtract.Amount1 = ""

        ExtractList.Add(newExtract)

        iSumAmount1 = 0


        LoadingMessage.Text = "This list takes a while to load - calculating Self Select Volumes "
        Me.Refresh()

        SetupSelfSelect(startdate, enddate, ExtractList, 0, 1, 0, 1)
        SetupSelfSelect(startdate, enddate, ExtractList, 1, 1, 0, 1)
        SetupSelfSelect(startdate, enddate, ExtractList, 2, 1, 0, 1)
        SetupSelfSelect(startdate, enddate, ExtractList, 0, 0, 0, 0)
        SetupSelfSelect(startdate, enddate, ExtractList, 1, 0, 0, 0)
        SetupSelfSelect(startdate, enddate, ExtractList, 2, 0, 0, 0)

        SetupSelfSelectTotals(ExtractList)
        SetupNewAccountTotals(ExtractList)
        SetupClosedAccounts(startdate, enddate, ExtractList)

        newExtract = New Extract
        newExtract.Description = "New Accounts Per Category "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)


        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Active Lender total "
        Me.Refresh()

        SetupNewVolumes(startdate, enddate, ExtractList, 5)
        SetupNewVolumes(startdate, enddate, ExtractList, 2)
        SetupNewVolumes(startdate, enddate, ExtractList, 3)
        SetupNewVolumes(startdate, enddate, ExtractList, 1)
        SetupNewVolumes(startdate, enddate, ExtractList, 4)

        SetupNewVolumesTotals(ExtractList)


        newExtract = New Extract
        newExtract.Description = "Smart Search Pass Rate  "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)


        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Active Lender total "
        Me.Refresh()

        SetupNewSSVolumes(startdate, enddate, ExtractList, 0)
        SetupNewSSVolumes(startdate, enddate, ExtractList, 1)

        SetupNewSSVolumesTotals(ExtractList)


        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "

        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)


        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumIUT = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Lender total "
        Me.Refresh()

        newExtract = New Extract
        newExtract.Description = "Active "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 0, 1, 0) 'each acc type for acive accounts, non-selective, mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 0, 1, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 0, 1, 0)
        SetupAccountTotalsTotals(ExtractList, 0, 1, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Idle "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 1, 1, 0) 'each acc type for idle accounts, non-selective, mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 1, 1, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 1, 1, 0)
        SetupAccountTotalsTotals(ExtractList, 1, 1, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Balance < £100 "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 2, 1, 0) 'each acc type for low balance accounts, non-selective, mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 2, 1, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 2, 1, 0)
        SetupAccountTotalsTotals(ExtractList, 2, 1, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Unfunded "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 3, 1, 0) 'each acc type for unfunded accounts, non-selective, mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 3, 1, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 3, 1, 0)
        SetupAccountTotalsTotals(ExtractList, 3, 1, 0)

        iSumAmount1 = 0

        newExtract = New Extract
        newExtract.Description = "Active "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 0, 0, 1) 'each acc type for acive accounts, selective, no mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 0, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 2, 0, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 0, 0, 0, 0) 'and then with mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 0, 0, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 0, 0, 0)
        SetupAccountTotalsTotals(ExtractList, 0, 0, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Idle "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 1, 0, 1) 'each acc type for idle accounts, selective, no-mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 1, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 2, 1, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 0, 1, 0, 0) 'and then with mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 1, 0, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 1, 0, 0)
        SetupAccountTotalsTotals(ExtractList, 1, 0, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Balance < £100 "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 2, 0, 1) 'each acc type for low balance accounts, selective, no-mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 2, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 2, 2, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 0, 2, 0, 0) 'and then with mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 2, 0, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 2, 0, 0)
        SetupAccountTotalsTotals(ExtractList, 2, 0, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Unfunded "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupAccountTotals(enddate, ExtractList, 0, 3, 0, 1) 'each acc type for unfunded accounts, selective, no-mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 3, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 2, 3, 0, 1)
        SetupAccountTotals(enddate, ExtractList, 0, 3, 0, 0) 'and then with mandate 
        SetupAccountTotals(enddate, ExtractList, 1, 3, 0, 0)
        SetupAccountTotals(enddate, ExtractList, 2, 3, 0, 0)
        SetupAccountTotalsTotals(ExtractList, 3, 0, 0)

        SetupAccountFinalTotal(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Deposits "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Lender total "
        Me.Refresh()

        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupDeposits(startdate, enddate, ExtractList, 0, 1, 0) 'each acc type for non-selective, mandate 
        SetupDeposits(startdate, enddate, ExtractList, 1, 1, 0)
        SetupDeposits(startdate, enddate, ExtractList, 2, 1, 0)
        SetupDepositsTotals(ExtractList, 1, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Self-Select Accounts "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupDeposits(startdate, enddate, ExtractList, 0, 0, 1) 'each acc type for selective, mandate 
        SetupDeposits(startdate, enddate, ExtractList, 1, 0, 1)
        SetupDeposits(startdate, enddate, ExtractList, 2, 0, 1)
        SetupDeposits(startdate, enddate, ExtractList, 0, 0, 0) 'and then with mandate 
        SetupDeposits(startdate, enddate, ExtractList, 1, 0, 0)
        SetupDeposits(startdate, enddate, ExtractList, 2, 0, 0)
        SetupDepositsTotals(ExtractList, 0, 1)

        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Withdrawals "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)


        LoadingMessage.Text = "This list takes a while to load - calculating Inactive Lender total "
        Me.Refresh()

        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupWithdrawals(startdate, enddate, ExtractList, 0, 1, 0) 'each acc type for non-selective, mandate 
        SetupWithdrawals(startdate, enddate, ExtractList, 1, 1, 0)
        SetupWithdrawals(startdate, enddate, ExtractList, 2, 1, 0)
        SetupWithdrawalsTotals(ExtractList, 1, 0)
        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Self-Select Accounts "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupWithdrawals(startdate, enddate, ExtractList, 0, 0, 1) 'each acc type for selective, mandate 
        SetupWithdrawals(startdate, enddate, ExtractList, 1, 0, 1)
        SetupWithdrawals(startdate, enddate, ExtractList, 2, 0, 1)
        SetupWithdrawals(startdate, enddate, ExtractList, 0, 0, 0) 'and then with mandate 
        SetupWithdrawals(startdate, enddate, ExtractList, 1, 0, 0)
        SetupWithdrawals(startdate, enddate, ExtractList, 2, 0, 0)
        SetupWithdrawalsTotals(ExtractList, 0, 1)

        SetupDepositWithdrawalTotals(ExtractList)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0
        iSumAmount4 = 0

        newExtract = New Extract
        newExtract.Description = "Client Account Free Balances "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Client Account Free Balances"
        Me.Refresh()
        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupFreeBalancesValues(enddate, ExtractList, 0, 1, 0) 'each acc type for non-selective, mandate 
        SetupFreeBalancesValues(enddate, ExtractList, 1, 1, 0)
        SetupFreeBalancesValues(enddate, ExtractList, 2, 1, 0)
        SetupFreeBalanceSubTotals(ExtractList, 1, 0)

        iSumAmount1 = 0
        newExtract = New Extract
        newExtract.Description = "Self-Select Accounts "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupFreeBalancesValues(enddate, ExtractList, 0, 0, 1) 'each acc type for selective, mandate 
        SetupFreeBalancesValues(enddate, ExtractList, 1, 0, 1)
        SetupFreeBalancesValues(enddate, ExtractList, 2, 0, 1)
        SetupFreeBalancesValues(enddate, ExtractList, 0, 0, 0) 'and then with mandate 
        SetupFreeBalancesValues(enddate, ExtractList, 1, 0, 0)
        SetupFreeBalancesValues(enddate, ExtractList, 2, 0, 0)
        SetupFreeBalanceSubTotals(ExtractList, 0, 1)

        SetupFreeBalanceTotals(ExtractList)

        newExtract = New Extract
        newExtract.Description = "Client Account Loan Balances "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)

        iSumAmount1 = 0
        iSumAmount2 = 0
        iSumAmount3 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating Client Account Loan Balances"
        Me.Refresh()
        newExtract = New Extract
        newExtract.Description = "Mandatory Mandates "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupLoanBalancesValues(enddate, ExtractList, 0, 1, 0) 'each acc type for non-selective, mandate 
        SetupLoanBalancesValues(enddate, ExtractList, 1, 1, 0)
        SetupLoanBalancesValues(enddate, ExtractList, 2, 1, 0)
        SetupLoanBalanceSubTotals(ExtractList, 1, 0)
        iSumAmount1 = 0

        newExtract = New Extract
        newExtract.Description = "Self-Select Accounts "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)
        SetupLoanBalancesValues(enddate, ExtractList, 0, 0, 1) 'each acc type for selective, mandate 
        SetupLoanBalancesValues(enddate, ExtractList, 1, 0, 1)
        SetupLoanBalancesValues(enddate, ExtractList, 2, 0, 1)
        SetupLoanBalancesValues(enddate, ExtractList, 0, 0, 0) 'and then with mandate 
        SetupLoanBalancesValues(enddate, ExtractList, 1, 0, 0)
        SetupLoanBalancesValues(enddate, ExtractList, 2, 0, 0)
        SetupLoanBalanceSubTotals(ExtractList, 0, 1)

        SetupLoanBalanceTotals(ExtractList)

        SetupAUMTotals(ExtractList)



        newExtract = New Extract
        newExtract.Description = "AUA Lender Categorisation "
        newExtract.Amount1 = ""
        ExtractList.Add(newExtract)


        iSumAmount1 = 0

        LoadingMessage.Text = "This list takes a while to load - calculating AUA totals "
        Me.Refresh()

        SetupAUMValues(enddate, ExtractList, 5)
        SetupAUMValues(enddate, ExtractList, 2)
        SetupAUMValues(enddate, ExtractList, 3)
        SetupAUMValues(enddate, ExtractList, 1)
        SetupAUMValues(enddate, ExtractList, 4)
        SetupAUMValues(enddate, ExtractList, 99)

        SetupAUMTotal(ExtractList)



        LoadingMessage.Text = "List complete"
        Me.Refresh()



        'finally ....

        DataGridView1.DataSource = ExtractList
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).Width = 80


        Dim senddate As String = ToDate.Value.ToString
        senddate = senddate.Split(" ")(0)
        Dim xenddate As String = Replace(senddate.ToString, "/", "")


        TextBox1.Text = "LenderMIDataControl" & xenddate
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)



    End Sub

    Private Sub SetupMandatoryMandates(startdate As Date, enddate As Date, Extractlist As List(Of Extract), accttype As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct m.userid, m.accountid
                   from mi_extract m
                    where  m.extractdate > @dt1
                    and  m.extractdate <= @dt2
                    and m.accounttype = @at1 
                    and m.activated = 5 and m.activated_bank = 5 and m.activated_cert = 5
                    and m.newaccount = 0
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
					and (m.clientClassification = 5 
                    or isnull(m.userTotalUnits,0) < @st1) 
                    order by m.accountid"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@st1", iSSThreshold))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                newExtract.Amount1 = ncount
                Select Case accttype
                    Case 0
                        newExtract.Description = "New Standard Account"
                        tbNMS.Text = newExtract.Amount1
                    Case 1
                        newExtract.Description = "New SIPP Account"
                        tbNMP.Text = newExtract.Amount1
                    Case 2
                        newExtract.Description = "New ISA Account"
                        tbNMI.Text = newExtract.Amount1
                End Select



                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupMandatoryMandatesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "Total"

        newExtract.Amount1 = iSumAmount1

        tbTNM.Text = iSumAmount1

        Extractlist.Add(newExtract)

        iSumAmount2 = iSumAmount1
        iSumAmount3 = iSumAmount1


    End Sub

    Private Sub SetupSelfSelect(startdate As Date, enddate As Date, Extractlist As List(Of Extract),
                                accttype As Integer, NewMandate As Integer, tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct m.userid, m.accountid
                   from mi_extract m
                    where  m.extractdate > @dt1
                    and  m.extractdate <= @dt2
                    and m.accounttype = @at1
                    and m.activated = 5 and m.activated_bank = 5 and m.activated_cert = 5
                    and m.newaccount = 0
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
					and (m.clientClassification < 5
                    and isnull(m.userTotalUnits,0) >= @st1) 
                    and m.mandatelender = @ml1
                    order by m.accountid"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@nm1", NewMandate))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                newExtract.Amount1 = ncount
                If NewMandate = 1 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "New Standard Account"
                            tbNSS.Text = newExtract.Amount1
                        Case 1
                            newExtract.Description = "New SIPP Account"
                            tbNSP.Text = newExtract.Amount1
                        Case 2
                            newExtract.Description = "New ISA Account"
                            tbNSI.Text = newExtract.Amount1
                    End Select
                Else
                    iSumAmount2 += newExtract.Amount1
                    Select Case accttype
                        Case 0
                            newExtract.Description = "New Standard Account with Mandate"
                            tbNSSM.Text = newExtract.Amount1

                        Case 1
                            newExtract.Description = "New SIPP Account with Mandate"
                            tbNSPM.Text = newExtract.Amount1

                        Case 2
                            newExtract.Description = "New ISA Account with Mandate"
                            tbNSIM.Text = newExtract.Amount1

                    End Select
                End If




                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupSelfSelectTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "Total"

        newExtract.Amount1 = iSumAmount1

        tbTNS.Text = iSumAmount1
        iSumAmount3 += iSumAmount1
        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupNewAccountTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract
        newExtract.Description = "Total New Accounts"
        newExtract.Amount1 = iSumAmount3
        tbTNA.Text = iSumAmount3
        Extractlist.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "Total New Mandates"
        newExtract.Amount1 = iSumAmount2
        tbTNMA.Text = iSumAmount2
        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupClosedAccounts(startdate As Date, enddate As Date, Extractlist As List(Of Extract))
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct m.userid, m.accountid
                   from mi_extract m
                    where m.extractdate > @dt1
                    and  m.extractdate <= @dt2
                    and m.closedaccount = 0  
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
                    order by m.accountid"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count

                Dim newExtract As New Extract
                newExtract.Amount1 = ncount
                newExtract.Description = "Total Number Of Closed Accounts"
                tbTCA.Text = newExtract.Amount1

                Extractlist.Add(newExtract)

            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupNewVolumes(startdate As Date, enddate As Date, Extractlist As List(Of Extract), classification As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct m.userid, m.accountid
                   from mi_extract m
                    where  m.extractdate > @dt1
                    and  m.extractdate <= @dt2
					and m.clientClassification = @at1
                    and m.activated = 5 and m.activated_bank = 5 and m.activated_cert = 5
					and m.newaccount = 0
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
                    order by m.accountid"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", classification))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                Select Case classification

                    Case 1
                        newExtract.Description = "Elective Professional"
                    Case 2
                        newExtract.Description = "Self-Certified Sophisticated"
                    Case 3
                        newExtract.Description = "High Net Worth"
                    Case 4
                        newExtract.Description = "Per-Se Professional"
                    Case 5
                        newExtract.Description = "Restricted Lender"
                End Select

                newExtract.Amount1 = ncount


                Select Case classification
                    Case 1
                        tbEPV.Text = newExtract.Amount1
                    Case 2
                        tbSSV.Text = newExtract.Amount1
                    Case 3
                        tbHNV.Text = newExtract.Amount1
                    Case 4
                        tbPSV.Text = newExtract.Amount1
                    Case 5
                        tbRLV.Text = newExtract.Amount1
                End Select


                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupNewVolumesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL NEW ACCOUNTS"

        newExtract.Amount1 = iSumAmount1


        tbNVT.Text = iSumAmount1

        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupNewSSVolumes(startdate As Date, enddate As Date, Extractlist As List(Of Extract), SmartSearchRes As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select distinct m.userid, m.accountid
                   from mi_extract m
                    where  m.extractdate > @dt1
                    and  m.extractdate <= @dt2
					and m.SmartSearchResToday = @at1
					and m.newaccount = 0
                    order by m.accountid"


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", SmartSearchRes))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                Dim newExtract As New Extract
                newExtract.Amount1 = ncount
                Select Case SmartSearchRes
                    Case 0
                        newExtract.Description = "First-Time Pass"
                        tbSSP.Text = newExtract.Amount1
                    Case 1
                        newExtract.Description = "Manual Intervention"
                        tbSSM.Text = newExtract.Amount1
                End Select


                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupNewSSVolumesTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "Total Lenders"

        newExtract.Amount1 = iSumAmount1

        tbSST.Text = iSumAmount1

        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupAccountTotals(enddate As Date, Extractlist As List(Of Extract), accttype As Integer,
                                  lender_status As Integer, tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim SQLX As String
                If mandatelender = 0 And tradinglender = 1 Then 'this is a mandatory mandate iteration
                    SQLX = " and (tradinglender = 1 or isnull(userTotalUnits,0) < @st1)  "
                Else 'this is a self select iteration
                    SQLX = " and (tradinglender = 0 and isnull(userTotalUnits,0) >= @st1) and mandatelender = @ml1 "
                End If

                Dim adapter As SqlDataAdapter = New SqlDataAdapter()

                MySQL = "select distinct  vt.accountid
                    from 
                        (select accountid, max(mi_extract_ID) as maxID
                          from mi_extract 
						  where extractdate <= @dt2
						  and accounttype = @at1
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
                          and activated = 5 and activated_bank = 5 
						  and lender_status = @am1 " & SQLX &
                         " group by accountid) vt 
						  INNER JOIN
                          mi_extract as t 
						  on t.accountid = vt.accountid 
						  and t.mi_extract_id = vt.maxID
                              "





                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@am1", lender_status))
                    .Add(New SqlParameter("@tl1", tradinglender))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count

                Dim newExtract As New Extract
                newExtract.Amount1 = ncount
                If tradinglender <> 0 And mandatelender = 0 Then
                    Select Case lender_status
                        Case 0
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbMAS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbMAP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbMAI.Text = newExtract.Amount1
                            End Select
                        Case 1
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbMIS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbMIP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbMII.Text = newExtract.Amount1
                            End Select
                        Case 2
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbMLS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbMLP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbMLI.Text = newExtract.Amount1
                            End Select
                        Case 3
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbMUS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbMUP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbMUI.Text = newExtract.Amount1
                            End Select
                    End Select
                End If

                If tradinglender = 0 And mandatelender = 1 Then
                    Select Case lender_status
                        Case 0
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSAS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSAP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSAI.Text = newExtract.Amount1
                            End Select
                        Case 1
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSIS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSIP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSII.Text = newExtract.Amount1
                            End Select
                        Case 2
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSLS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSLP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSLI.Text = newExtract.Amount1
                            End Select
                        Case 3
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSUS.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSUP.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSUI.Text = newExtract.Amount1
                            End Select
                    End Select
                End If
                If tradinglender = 0 And mandatelender = 0 Then
                    Select Case lender_status
                        Case 0
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSASM.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSAPM.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSAIM.Text = newExtract.Amount1
                            End Select
                        Case 1
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSISM.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSIPM.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSIIM.Text = newExtract.Amount1
                            End Select
                        Case 2
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSLSM.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSLPM.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSLIM.Text = newExtract.Amount1
                            End Select
                        Case 3
                            Select Case accttype
                                Case 0
                                    newExtract.Description = "Standard Account"
                                    tbSUSM.Text = newExtract.Amount1
                                Case 1
                                    newExtract.Description = "SIPP Account"
                                    tbSUPM.Text = newExtract.Amount1
                                Case 2
                                    newExtract.Description = "ISA Account"
                                    tbSUIM.Text = newExtract.Amount1
                            End Select
                    End Select
                End If





                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub

    Private Sub SetupAccountTotalsTotals(Extractlist As List(Of Extract), lender_status As Integer,
                                         tradinglender As Integer, mandatelender As Integer)

        Dim newExtract As New Extract
        If tradinglender <> 0 Then
            Select Case lender_status
                Case 0
                    newExtract.Description = "TOTAL ACTIVE ACCOUNTS"
                    tbMAT.Text = iSumAmount1
                Case 1
                    newExtract.Description = "TOTAL IDLE ACCOUNTS"
                    tbMIT.Text = iSumAmount1
                Case 2
                    newExtract.Description = "TOTAL BALANCE < £100 ACCOUNTS"
                    tbMLT.Text = iSumAmount1
                Case 3
                    newExtract.Description = "TOTAL UNFUNDED ACCOUNTS"
                    tbMUT.Text = iSumAmount1

            End Select
        End If

        If tradinglender = 0 Then
            Select Case lender_status
                Case 0
                    newExtract.Description = "TOTAL ACTIVE ACCOUNTS"
                    tbSAT.Text = iSumAmount1
                Case 1
                    newExtract.Description = "TOTAL IDLE ACCOUNTS"
                    tbSIT.Text = iSumAmount1
                Case 2
                    newExtract.Description = "TOTAL BALANCE < £100 ACCOUNTS"
                    tbSLT.Text = iSumAmount1
                Case 3
                    newExtract.Description = "TOTAL UNFUNDED ACCOUNTS"
                    tbSUT.Text = iSumAmount1

            End Select
        End If



        newExtract.Amount1 = iSumAmount1

        iSumAmount2 += iSumAmount1


        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupAccountFinalTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL  ACCOUNTS"
        tbTAT.Text = iSumAmount2

        newExtract.Amount1 = iSumAmount2

        Extractlist.Add(newExtract)

    End Sub

    Private Sub SetupDeposits(startdate As Date, enddate As Date, Extractlist As List(Of Extract), accttype As Integer,
                                  tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Dim nSumm As Integer
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim SQLX As String
                If mandatelender = 0 And tradinglender = 1 Then 'this is a mandatory mandate iteration
                    SQLX = " and (tradinglender = 1 or isnull(userTotalUnits,0) < @st1)  "
                Else 'this is a selef select iteration
                    SQLX = " and (tradinglender = 0 and isnull(userTotalUnits,0) >= @st1) and mandatelender = @ml1  "
                End If

                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select isnull(sum(depositstoday),0) as theamount
                          from mi_extract 
						  where extractdate <= @dt2
                          and extractdate > @dt1 
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
						  and accounttype = @at1 " & SQLX

                'and tradinglender = @tl1
                'and mandatelender = @ml1 
                '    "


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@tl1", tradinglender))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count

                If ds1.Tables(0).Rows.Count > 0 Then
                    With ds1.Tables(0).Rows(0)
                        nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
                    End With
                Else
                    nSumm = 0
                End If


                Dim newExtract As New Extract
                newExtract.Amount1 = nSumm
                If tradinglender <> 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbDMS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbDMP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbDMI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                If tradinglender = 0 And mandatelender = 1 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbDSS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbDSP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbDSI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If
                If tradinglender = 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbDSSM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbDSPM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbDSIM.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If



                iSumAmount1 += newExtract.Amount1

                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupDepositsTotals(Extractlist As List(Of Extract), tradinglender As Integer, mandatelender As Integer)

        Dim newExtract As New Extract
        newExtract.Description = "Total"
        If tradinglender = 1 Then
            tbDMT.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If

        If tradinglender = 0 Then
            tbDST.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If

        newExtract.Amount1 = iSumAmount1
        Extractlist.Add(newExtract)
        iSumAmount2 += iSumAmount1


    End Sub

    Private Sub SetupWithdrawals(startdate As Date, enddate As Date, Extractlist As List(Of Extract), accttype As Integer,
                                  tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Dim nSumm As Integer
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim SQLX As String
                If mandatelender = 0 And tradinglender = 1 Then 'this is a mandatory mandate iteration
                    SQLX = " and (tradinglender = 1 or isnull(userTotalUnits,0) < @st1)  "
                Else 'this is a selef select iteration
                    SQLX = " and (tradinglender = 0 and isnull(userTotalUnits,0) >= @st1) and mandatelender = @ml1 "
                End If

                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                MySQL = "select isnull(sum(withdrawalstoday),0) as theamount
                          from mi_extract 
						  where extractdate <= @dt2
                          and extractdate > @dt1 
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
						  and accounttype = @at1 " & SQLX
                'and tradinglender = @tl1
                'and mandatelender = @ml1 
                '    "


                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@dt1", startdate))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@tl1", tradinglender))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count


                If ds1.Tables(0).Rows.Count > 0 Then
                    With ds1.Tables(0).Rows(0)
                        nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
                    End With
                Else
                    nSumm = 0
                End If

                Dim newExtract As New Extract
                newExtract.Amount1 = nSumm
                If tradinglender <> 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbWMS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbWMP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbWMI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                If tradinglender = 0 And mandatelender = 1 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbWSS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbWSP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbWSI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If
                If tradinglender = 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbWSSM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbWSPM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbWSIM.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If



                iSumAmount1 += newExtract.Amount1

                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using

    End Sub

    Private Sub SetupWithdrawalsTotals(Extractlist As List(Of Extract), tradinglender As Integer, mandatelender As Integer)

        Dim newExtract As New Extract
        newExtract.Description = "Total"
        If tradinglender = 1 Then
            tbWMT.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If

        If tradinglender = 0 Then
            tbWST.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If
        newExtract.Amount1 = iSumAmount1

        Extractlist.Add(newExtract)
        iSumAmount3 += iSumAmount1


    End Sub

    Private Sub SetupDepositWithdrawalTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract
        newExtract.Description = "TOTAL DEPOSITS"
        newExtract.Amount1 = iSumAmount2
        tbTDT.Text = PenceToCurrencyStringPounds(iSumAmount2)
        Extractlist.Add(newExtract)

        newExtract = New Extract
        newExtract.Description = "TOTAL WITHDRAWALS"
        newExtract.Amount1 = iSumAmount3
        tbTWT.Text = PenceToCurrencyStringPounds(iSumAmount3)
        Extractlist.Add(newExtract)


    End Sub

    Private Sub SetupFreeBalancesValues(enddate As Date, Extractlist As List(Of Extract), accttype As Integer,
                                        tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String
        Dim dAmount As Double

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim SQLX As String
                If mandatelender = 0 And tradinglender = 1 Then 'this is a mandatory mandate iteration
                    SQLX = " and (tradinglender = 1 or isnull(userTotalUnits,0) < @st1)  "
                Else 'this is a selef select iteration
                    SQLX = " and (tradinglender = 0 and isnull(userTotalUnits,0) >= @st1) and mandatelender = @ml1 "
                End If

                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                '        MySQL = "select isnull(sum(f.theamount),0) as theamount
                '        from (
                '         select isnull(cashbalance,0) as theamount from GET_MI_EXTRACT 
                '              where extractdate <= @dt2
                '                and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
                'and accounttype = @at1 " & SQLX &
                '                  " ) f "

                MySQL = " select isnull(sum(f.theamount),0) as theamount
                from (
                 select isnull(cashbalance,0) as theamount, accountid, mi_extract_id from MI_EXTRACT 
                      where extractdate <= @dt2
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
						  and accounttype = @at1 " & SQLX &
                          " group by accountid, mi_extract_id , cashbalance) f INNER JOIN
                 (SELECT   accountid, MAX(MI_EXTRACT_ID) AS maxID
                 FROM      dbo.MI_EXTRACT
				 where extractdate <= @dt2
				 group by accountid ) g  ON f.accountid = g.accountid AND f.MI_EXTRACT_ID = g.maxID"

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@tl1", tradinglender))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
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
                Else
                    nSumm = 0
                End If

                Dim newExtract As New Extract
                newExtract.Amount1 = nSumm
                If tradinglender <> 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbFMS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbFMP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbFMI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                If tradinglender = 0 And mandatelender = 1 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbFSS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbFSP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbFSI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If
                If tradinglender = 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbFSSM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbFSPM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbFSIM.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                iSumAmount1 += nSumm


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub

    Private Sub SetupFreeBalanceSubTotals(Extractlist As List(Of Extract), tradinglender As Integer, mandatelender As Integer)

        Dim newExtract As New Extract
        newExtract.Description = "Total"
        If tradinglender = 1 Then
            tbFMT.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If

        If tradinglender = 0 Then
            tbFST.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If
        newExtract.Amount1 = iSumAmount1
        ' tbTNM.Text = PenceToCurrencyStringPounds(iSumAmount1)
        Extractlist.Add(newExtract)
        iSumAmount3 += iSumAmount1


    End Sub

    Private Sub SetupFreeBalanceTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL CLIENT ACCOUNT FREE BALANCES"

        newExtract.Amount1 = iSumAmount3

        tbFTT.Text = PenceToCurrencyStringPounds(iSumAmount3)

        iSumAmount4 += iSumAmount3

        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupLoanBalancesValues(enddate As Date, Extractlist As List(Of Extract), accttype As Integer,
                                        tradinglender As Integer, mandatelender As Integer)
        Dim MySQL, strConn As String
        Dim dAmount As Double

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim SQLX As String
                If mandatelender = 0 And tradinglender = 1 Then 'this is a mandatory mandate iteration
                    SQLX = " and (tradinglender = 1 or isnull(userTotalUnits,0) < @st1)  "
                Else 'this is a selef select iteration
                    SQLX = " and (tradinglender = 0 and isnull(userTotalUnits,0) >= @st1) and mandatelender = @ml1 "
                End If

                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                '        MySQL = "select isnull(sum(f.theamount),0) as theamount
                '        from (
                '         select isnull(loanbalance,0) as theamount from GET_MI_EXTRACT 
                '              where extractdate <= @dt2
                '                and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
                'and accounttype = @at1  " & SQLX &
                '                  " ) f  "


                MySQL = " select isnull(sum(f.theamount),0) as theamount
                from (
                 select isnull(loanbalance,0) as theamount, accountid, mi_extract_id from MI_EXTRACT 
                      where extractdate <= @dt2
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
						  and accounttype = @at1 " & SQLX &
                          " group by accountid, mi_extract_id , loanbalance) f INNER JOIN
                 (SELECT   accountid, MAX(MI_EXTRACT_ID) AS maxID
                 FROM      dbo.MI_EXTRACT
				 where extractdate <= @dt2
				 group by accountid ) g  ON f.accountid = g.accountid AND f.MI_EXTRACT_ID = g.maxID"

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters
                    .Add(New SqlParameter("@at1", accttype))
                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@tl1", tradinglender))
                    .Add(New SqlParameter("@ml1", mandatelender))
                    .Add(New SqlParameter("@st1", iSSThreshold))
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
                newExtract.Amount1 = nSumm
                If tradinglender <> 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbLMS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbLMP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbLMI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                If tradinglender = 0 And mandatelender = 1 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbLSS.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbLSP.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbLSI.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If
                If tradinglender = 0 And mandatelender = 0 Then
                    Select Case accttype
                        Case 0
                            newExtract.Description = "Standard Account"
                            tbLSSM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 1
                            newExtract.Description = "SIPP Account"
                            tbLSPM.Text = PenceToCurrencyStringPounds(nSumm)
                        Case 2
                            newExtract.Description = "ISA Account"
                            tbLSIM.Text = PenceToCurrencyStringPounds(nSumm)
                    End Select
                End If

                iSumAmount1 += nSumm


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub

    Private Sub SetupLoanBalanceSubTotals(Extractlist As List(Of Extract), tradinglender As Integer, mandatelender As Integer)

        Dim newExtract As New Extract
        newExtract.Description = "Total"
        If tradinglender = 1 Then
            tbLMT.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If

        If tradinglender = 0 Then
            tbLST.Text = PenceToCurrencyStringPounds(iSumAmount1)
        End If
        newExtract.Amount1 = iSumAmount1
        ' tbTNM.Text = iSumAmount1
        Extractlist.Add(newExtract)
        iSumAmount3 += iSumAmount1


    End Sub

    Private Sub SetupLoanBalanceTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL LENDER LOAN BALANCES"

        newExtract.Amount1 = iSumAmount3

        tbLTL.Text = PenceToCurrencyStringPounds(iSumAmount3)

        iSumAmount2 += iSumAmount3

        Extractlist.Add(newExtract)



    End Sub



    Private Sub SetupAUMTotals(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL AUA (LOANS + CASH)"

        newExtract.Amount1 = iSumAmount2 + iSumAmount4

        tbTLC.Text = PenceToCurrencyStringPounds(iSumAmount2 + iSumAmount4)


        Extractlist.Add(newExtract)



    End Sub

    Private Sub SetupAUMValues(enddate As Date, Extractlist As List(Of Extract),
                                        clientclassification As Integer)
        Dim MySQL, strConn As String

        Dim dr1, dr2, dr3 As DataRow
        Dim ds1 As DataSet
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                '        MySQL = "select isnull(sum(f.loanbalance + f.cashbalance),0) as theamount
                '        from (
                '         select isnull(cashbalance,0) as cashbalance, isnull(loanbalance, 0) as loanbalance 
                '             from GET_MI_EXTRACT 
                '              where extractdate <= @dt2
                '                and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)      
                'and isnull(clientclassification,9) = @at1) f"

                MySQL = " select isnull(sum(f.loanbalance + f.cashbalance),0) as theamount
                from (
                 select isnull(loanbalance,0) as loanbalance, isnull(cashbalance,0) as cashbalance, 
                      accountid, mi_extract_id from MI_EXTRACT 
                      where extractdate <= @dt2
                        and accountid not in (3163, 3709, 3710, 3711, 3712, 3713)   
						  and isnull(clientclassification,9) = @at1 
                         group by accountid, mi_extract_id , loanbalance, cashbalance) f 
                  INNER JOIN
                 (SELECT   accountid, MAX(MI_EXTRACT_ID) AS maxID
                 FROM      dbo.MI_EXTRACT
				 where extractdate <= @dt2
				 group by accountid ) g  ON f.accountid = g.accountid AND f.MI_EXTRACT_ID = g.maxID"

                Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                con.Open()
                cmd.Parameters.Clear()
                With cmd.Parameters

                    .Add(New SqlParameter("@dt2", enddate))
                    .Add(New SqlParameter("@at1", clientclassification))
                End With
                adapter.SelectCommand = cmd

                ds1 = New DataSet

                adapter.Fill(ds1)

                Dim ncount As Integer = ds1.Tables(0).Rows.Count

                Dim nSumm As Double = 0
                If ds1.Tables(0).Rows.Count > 0 Then
                    With ds1.Tables(0).Rows(0)
                        nSumm = ds1.Tables(0).Rows(0).Item(“theamount”)
                    End With
                End If

                Dim newExtract As New Extract
                Select Case clientclassification

                    Case 1
                        newExtract.Description = "Elective Professional"
                    Case 2
                        newExtract.Description = "Self-Certified Sophisticated"
                    Case 3
                        newExtract.Description = "High Net Worth"
                    Case 4
                        newExtract.Description = "Per-Se Professional"
                    Case 5
                        newExtract.Description = "Restricted Lender"
                    Case 99
                        newExtract.Description = "Categorisation Pending"
                End Select

                newExtract.Amount1 = nSumm


                Select Case clientclassification
                    Case 1
                        tbTAE.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                    Case 2
                        tbTAS.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                    Case 3
                        tbTAH.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                    Case 4
                        tbTAP.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                    Case 5
                        tbTAR.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                    Case 99
                        tbTCP.Text = PenceToCurrencyStringPounds(newExtract.Amount1)
                End Select


                iSumAmount1 += newExtract.Amount1


                Extractlist.Add(newExtract)
            Catch ex As Exception
            Finally

            End Try
        End Using
    End Sub



    Private Sub SetupAUMTotal(Extractlist As List(Of Extract))

        Dim newExtract As New Extract

        newExtract.Description = "TOTAL AUA"

        newExtract.Amount1 = iSumAmount1

        tbTAU.Text = PenceToCurrencyStringPounds(iSumAmount1)


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

    Public Shared Function CurrencyStringPoundsToPence(ByVal sField) As Integer
        Dim d As Double
        Dim i, rVal As Integer
        Dim s As String

        If IsDBNull(sField) Then
            rVal = 0
        Else
            Try
                s = sField.ToString
                i = InStr(s, "£")
                If i > 0 Then
                    s = s.Remove(i - 1, 1)
                    Trim(s)
                End If

                If Double.TryParse(s, d) Then
                    rVal = CInt(d * 100)
                Else
                    rVal = 0
                End If
            Catch ex As Exception
                rVal = 0
            End Try
        End If
        CurrencyStringPoundsToPence = rVal
    End Function
End Class

