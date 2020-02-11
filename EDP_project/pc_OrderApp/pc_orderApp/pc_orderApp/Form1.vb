Imports System.Console
Imports System.Data.OleDb

Public Class frm_pcOrder
    Dim cost As Decimal = 0
    Dim shoppinglist(13) As String

    Dim con As New OleDb.OleDbConnection
    Dim dbProvide As String
    Dim dbSource As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim sql As String
    Dim maxRows As Integer
    Dim inc As Integer = -1

    Dim conOrder As New OleDb.OleDbConnection
    Dim dbProvideOrder As String
    Dim dbSourceOrder As String
    Dim dsOrder As New DataSet
    Dim daOrder As OleDb.OleDbDataAdapter
    Dim sqlOrder As String
    Dim maxRowsOrder As Integer
    Dim incOrder As Integer = -1
   

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        MsgBox("Please make sure to keep the folder EDP_project in the 'C' drive")
        'specify type of database being used
        dbProvide = "PROVIDER= Microsoft.jet.OLEDB.4.0;"
        'specify source db location
        dbSource = "Data Source = C:\EDP_project\pc_OrderApp\pc_orderApp\StaffLogin.mdb"

        'use connection object to create a connection string
        con.ConnectionString = dbProvide & dbSource

        'open connection to db
        con.Open()

        sql = "SELECT * FROM tblStaff"
        'use adapter to run sql statement on connection
        da = New OleDb.OleDbDataAdapter(sql, con)
        'use adapter to put result in dataset
        da.Fill(ds, "StaffLogin")

        maxRows = ds.Tables("StaffLogin").Rows.Count
        MsgBox("Connection successfull")

        'close connection to db
        con.Close()





        'REMOVING ALL TABS EXCEPT login
        tab_pcOrder.TabPages.Remove(tab_customer_details)
        tab_pcOrder.TabPages.Remove(tab_orders)
        tab_pcOrder.TabPages.Remove(tab_payment)
        tab_pcOrder.TabPages.Remove(tab_breakdown)
        tab_pcOrder.TabPages.Remove(tab_administrator)
        tab_pcOrder.TabPages.Remove(Manager_TAB)

        lbl_timeDate.Text = DateTime.Now.ToLongDateString()



    End Sub

  
    Private Sub databseconnection()
       
        'specify type of database being used
        dbProvideOrder = "PROVIDER= Microsoft.jet.OLEDB.4.0;"
        'specify source db location
        dbSourceOrder = "Data Source = C:\EDP_project\pc_OrderApp\pc_orderApp\Orders.mdb"

        'use connection object to create a connection string
        conOrder.ConnectionString = dbProvideOrder & dbSourceOrder

        'open connection to db
        conOrder.Open()

        sqlOrder = "SELECT * FROM tblOrders"
        'use adapter to run sql statement on connection
        daOrder = New OleDb.OleDbDataAdapter(sqlOrder, conOrder)
        'use adapter to put result in dataset
        daOrder.Fill(dsOrder, "Orders")

        maxRowsOrder = dsOrder.Tables("Orders").Rows.Count
        '  MsgBox("Connection successfull")

        'close connection to db
        conOrder.Close()
       
    End Sub

    'LOGIN BUTTON CLICK EVENTS
    Private Sub btn_login_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_login.Click

        'VERYFYING DETAILS ENTERED WITH IF STATEMENT
        For x = 0 To maxRows - 1
            If tbx_staffid.Text = ds.Tables("StaffLogin").Rows(x).Item(1) And tbx_pass.Text = ds.Tables("StaffLogin").Rows(x).Item(2) Then
                MsgBox("Access Guaranted")
                'IF DETAILS ARE CORRECT THEN ALL TABS WILL BE DISPLAYED AND LOGIN WILL BE REMOVED
                tab_pcOrder.TabPages.Remove(tab_login)
                tab_pcOrder.TabPages.Add(tab_customer_details)
                tab_pcOrder.TabPages.Add(tab_orders)
                tab_pcOrder.TabPages.Add(tab_payment)
                tab_pcOrder.TabPages.Add(tab_breakdown)

                'SELECTING THE NEXT TAB AND DISPLAYING THE STAFF ID
                For Each ctl As Control In tab_customer_details.Controls
                    If TypeOf ctl Is TextBox Then
                        DirectCast(ctl, TextBox).Clear()
                    End If
                Next

                tab_pcOrder.TabPages(0).Enabled = True
                tab_pcOrder.SelectedTab = tab_customer_details
                tbx_tab2staffID.Text = tbx_staffid.Text

                'UNTIL CUSTOMER DETAILS ARE NOT ENTERED ALL THE OTHER TABS WILL BE DISABLED
                Me.tab_pcOrder.TabPages(1).Enabled = False
                Me.tab_pcOrder.TabPages(2).Enabled = False
                Me.tab_pcOrder.TabPages(3).Enabled = False

                'IF THE CONDITION IS TRUE THE METHOD WILL END
                Exit Sub

            End If

        Next

        'IF WRONG DETAILS HAVE BEEN ENTERED THEN A MESSAGE BOX WILL BE DISPLAYED
        MsgBox("Access Denied")

        For Each ctl As Control In tab_login.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next

    End Sub

    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click

        End

    End Sub

    'LOGOUT BUTTON EVENTS
    Private Sub btn_logout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_logout.Click

        'REMOVING ALL THE TABS AND GOING BACK TO login TAB
        tab_pcOrder.TabPages.Remove(tab_customer_details)
        tab_pcOrder.TabPages.Remove(tab_orders)
        tab_pcOrder.TabPages.Remove(tab_payment)
        tab_pcOrder.TabPages.Remove(tab_breakdown)

        'CLEARING THE TEXTBOXES INN THE login TAB
        tab_pcOrder.TabPages.Add(tab_login)
        For Each ctl As Control In tab_login.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
    End Sub

    'CUSTOMER DETAILS next BUTTON EVENST
    Private Sub btn_next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_next.Click



        'CHECKING FOR VALUES IN ALL THE TEXTBOXES IN THE CUSTOMER DETAILS tab
        If tbx_firstname.Text <> "" And tbx_surname.Text <> "" And tbx_street.Text <> "" And tbx_houseno.Text <> "" And tbx_county.Text <> "" And tbx_postcode.Text <> "" And tbx_telephone.Text <> "" Then
            If IsNumeric(tbx_houseno.Text) Then
                If IsNumeric(tbx_telephone.Text) Then
                    For Each ctl As Control In tab_orders.Controls
                        If TypeOf ctl Is ComboBox Then
                            DirectCast(ctl, ComboBox).SelectedIndex() = -1
                        End If
                    Next
                    For Each ctl As Control In tab_orders.Controls
                        If TypeOf ctl Is RadioButton Then
                            DirectCast(ctl, RadioButton).Checked = False
                        End If
                    Next

                    '--------------------------DATABASE CONNECTION PASTING DETAILS IN DATABASE (ORDERS)------------------------------------
                    databseconnection()
                    Dim cbOrders As New OleDb.OleDbCommandBuilder(daOrder)
                    Dim dsNewRowOrder As DataRow

                    dsNewRowOrder = dsOrder.Tables("Orders").NewRow()
                    dsNewRowOrder.Item(2) = tbx_firstname.Text & " " & tbx_surname.Text
                    dsNewRowOrder.Item(3) = tbx_street.Text
                    dsNewRowOrder.Item(4) = tbx_county.Text
                    dsNewRowOrder.Item(5) = tbx_postcode.Text
                    dsNewRowOrder.Item(6) = tbx_telephone.Text

                    dsOrder.Tables("Orders").Rows.Add(dsNewRowOrder)
                    daOrder.Update(dsOrder, "Orders")
                    ' MsgBox("Copied")

                    '---------------------DATABASE Connection end------------------
                    Me.tab_pcOrder.TabPages(1).Enabled = True
                    tab_pcOrder.SelectedTab = tab_orders
                    tbx_tab3staffID.Text = tbx_staffid.Text

                    Me.tab_pcOrder.TabPages(0).Enabled = False
                Else
                    MsgBox("Invalid Value in Telephone Field")
                End If
                Else
                    MsgBox("Is that a House Number???")
                End If




        Else
            MsgBox("Please complete the required fields")
        End If

    End Sub


    '...................ORDERS TAB EVENTS START..................
    Private Sub btn_tabOrder_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_tabOrder_Next.Click


        If cbx_case.SelectedIndex = 1 Then
            shoppinglist(0) = cbx_case.Text
            cost = cost + 77.98
        ElseIf cbx_case.SelectedIndex = 2 Then
            shoppinglist(0) = cbx_case.Text
            cost = cost + 22.49
        ElseIf cbx_case.SelectedIndex = 3 Then
            shoppinglist(0) = cbx_case.Text
            cost = cost + 49.58
        Else
            shoppinglist(0) = "CASE not selected"
        End If

        If cbx_motherboard.SelectedIndex = 1 Then
            shoppinglist(1) = cbx_motherboard.Text
            cost = cost + 89.99
        ElseIf cbx_motherboard.SelectedIndex = 2 Then
            shoppinglist(1) = cbx_motherboard.Text
            cost = cost + 204.99
        ElseIf cbx_motherboard.SelectedIndex = 3 Then
            shoppinglist(1) = cbx_motherboard.Text
            cost = cost + 79.99
        Else
            shoppinglist(1) = "MOTHERBOARD not selected"
        End If


        If cbx_powersupply.SelectedIndex = 1 Then
            shoppinglist(2) = cbx_powersupply.Text
            cost = cost + 89.99
        ElseIf cbx_powersupply.SelectedIndex = 2 Then
            shoppinglist(2) = cbx_powersupply.Text
            cost = cost + 204.99
        ElseIf cbx_powersupply.SelectedIndex = 3 Then
            shoppinglist(2) = cbx_powersupply.Text
            cost = cost + 79.99
        Else
            shoppinglist(2) = "POWER SUPPLY not selected"
        End If

        If cbx_netcard.SelectedIndex = 1 Then
            shoppinglist(3) = cbx_netcard.Text
            cost = cost + 19.99
        ElseIf cbx_netcard.SelectedIndex = 2 Then
            shoppinglist(3) = cbx_netcard.Text
            cost = cost + 34.99
        ElseIf cbx_netcard.SelectedIndex = 3 Then
            shoppinglist(3) = cbx_netcard.Text
            cost = cost + 29.99
        Else
            shoppinglist(3) = "NETWORK CARD not selected"
        End If

        If cbx_graphcard.SelectedIndex = 1 Then
            shoppinglist(4) = cbx_graphcard.Text
            cost = cost + 74.99
        ElseIf cbx_graphcard.SelectedIndex = 2 Then
            shoppinglist(4) = cbx_graphcard.Text
            cost = cost + 119.99
        ElseIf cbx_graphcard.SelectedIndex = 3 Then
            shoppinglist(4) = cbx_graphcard.Text
            cost = cost + 149.99
        Else
            shoppinglist(4) = "GRAPHIC CARD not selected"
        End If

        If cbx_sndcard.SelectedIndex = 1 Then
            shoppinglist(5) = cbx_sndcard.Text
            cost = cost + 39.99
        ElseIf cbx_sndcard.SelectedIndex = 2 Then
            shoppinglist(5) = cbx_sndcard.Text
            cost = cost + 49.99
        ElseIf cbx_sndcard.SelectedIndex = 3 Then
            shoppinglist(5) = cbx_sndcard.Text
            cost = cost + 27.99
        Else
            shoppinglist(5) = "SOUND CARD not selected"
        End If

        'OPTICAL DRIVE MULTIPLE CHOICES STRAT
        If chkbx_OD1.Checked = True Then
            shoppinglist(6) = chkbx_OD1.Text
            cost = cost + 19.99
        Else
            shoppinglist(6) = ""
        End If

        If chkbx_OD2.Checked = True Then
            shoppinglist(7) = chkbx_OD2.Text
            cost = cost + 24.99
        Else
            shoppinglist(7) = ""
        End If

        If chkbx_OD3.Checked = True Then
            shoppinglist(8) = chkbx_OD3.Text
            cost = cost + 29.99
        Else
            shoppinglist(8) = ""
        End If
        'OPTRICAL DRIVES MULTIPLE CHOICES END


        'HARD DRIVE MULTIPLE CHOICES START
        If chkbx_HHD1.Checked = True Then
            shoppinglist(9) = chkbx_HHD1.Text
            cost = cost + 64.99
        Else
            shoppinglist(9) = ""
        End If

        If chkbx_HHD2.Checked = True Then
            shoppinglist(10) = chkbx_HHD2.Text
            cost = cost + 129.99
        Else
            shoppinglist(10) = ""
        End If

        If chkbx_HHD3.Checked = True Then
            shoppinglist(11) = chkbx_HHD3.Text
            cost = cost + 99.99
        Else
            shoppinglist(11) = ""
        End If
        'HARD DRIVES MULTIPLE CHOICES END

        If cbx_ram.SelectedIndex = 1 Then
            shoppinglist(12) = cbx_ram.Text
            cost = cost + 69.99
        ElseIf cbx_ram.SelectedIndex = 2 Then
            shoppinglist(12) = cbx_ram.Text
            cost = cost + 29.99
        ElseIf cbx_ram.SelectedIndex = 3 Then
            shoppinglist(12) = cbx_ram.Text
            cost = cost + 69.99
        Else
            shoppinglist(12) = "RAM not selected"
        End If

        If cbx_cpu.SelectedIndex = 1 Then
            shoppinglist(13) = cbx_cpu.Text
            cost = cost + 179.99
        ElseIf cbx_cpu.SelectedIndex = 2 Then
            shoppinglist(13) = cbx_cpu.Text
            cost = cost + 174.99
        ElseIf cbx_cpu.SelectedIndex = 3 Then
            shoppinglist(13) = cbx_cpu.Text
            cost = cost + 99.99
        Else
            shoppinglist(13) = "CPU not selected"
        End If

        MsgBox("£ " & cost)

        For Each ctl As Control In tab_payment.Controls
            If TypeOf ctl Is RadioButton Then
                DirectCast(ctl, RadioButton).Checked = False
            End If
        Next
        For Each ctl As Control In tab_payment.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
        If cost > 0 Then
            tab_pcOrder.TabPages(2).Enabled = True
            tab_pcOrder.SelectedTab = tab_payment
            tbxPayement_staffID.Text = tbx_staffid.Text
            Me.tab_pcOrder.TabPages(1).Enabled = False
        Else
            MsgBox("Nothing has been purchased")
        End If
    End Sub
    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click

        For Each ctl As Control In tab_customer_details.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
        Me.tab_pcOrder.TabPages(0).Enabled = True
        tab_pcOrder.SelectedTab = tab_customer_details
        tbx_tab2staffID.Text = tbx_staffid.Text

        Me.tab_pcOrder.TabPages(1).Enabled = False
        Me.tab_pcOrder.TabPages(2).Enabled = False
        Me.tab_pcOrder.TabPages(3).Enabled = False

    End Sub
    '................ORDERS EVENTS END................

    '.................PAYMENT tab EVENTS START..................
    Private Sub btnradio_payment_1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnradio_payment_1.CheckedChanged
        tbx_deposit.Text = "Not Applicable"
        tbx_interest.Text = "Not Applicable"
        tbx_monthlypayment.Text = "Not Applicable"
        tbx_totalcost.Text = "£ " & Math.Round(cost, 2)

        lblSummary_mnthpayment.Text = "Monthly Payment"

    End Sub

    Private Sub btnradio_payment_2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnradio_payment_2.CheckedChanged
        Dim depositCost As Decimal = cost / 100 * 10
        Dim monthlyPayment As Decimal = (cost - depositCost) / 6

        tbx_deposit.Text = "£ " & Math.Round(depositCost, 2)
        tbx_monthlypayment.Text = "£ " & Math.Round(monthlyPayment, 2)
        tbx_interest.Text = "Interest Free"
        tbx_totalcost.Text = "£ " & Math.Round(cost, 2)
        lblSummary_mnthpayment.Text = "6 Monthly Payment"


    End Sub

    Private Sub btnradio_payment_3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnradio_payment_3.CheckedChanged
        Dim depositCost As Decimal = Math.Round(cost / 100 * 10, 2)
        Dim remainingAmount As Decimal = Math.Round(cost - depositCost, 2)
        Dim interest As Decimal = Math.Round(remainingAmount / 100 * 13, 2)
        Dim monthlyPayment As Decimal = Math.Round((remainingAmount + interest) / 12, 2)
        Dim totalcost As Decimal = Math.Round(cost + interest, 2)

        tbx_deposit.Text = "£ " & (depositCost)
        tbx_monthlypayment.Text = "£ " & (monthlyPayment)
        tbx_interest.Text = "£ " & (interest)
        tbx_totalcost.Text = "£ " & (totalcost)
        lblSummary_mnthpayment.Text = "12 Monthly Payment"

    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click


        Me.tab_pcOrder.TabPages(1).Enabled = True
        tab_pcOrder.SelectedTab = tab_orders
        Me.tab_pcOrder.TabPages(2).Enabled = False

        cost = 0

    End Sub


    Private Sub btn_payment_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_payment_cancel.Click

        cost = 0

        'CLEARING VALUES THE payment TAB
        For Each ctl As Control In tab_customer_details.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
        Me.tab_pcOrder.TabPages(0).Enabled = True
        tab_pcOrder.SelectedTab = tab_customer_details
        tbx_tab2staffID.Text = tbx_staffid.Text

        Me.tab_pcOrder.TabPages(1).Enabled = False
        Me.tab_pcOrder.TabPages(2).Enabled = False
    End Sub
    '.............END PAYMENT SECTION EVENTS................

    Private Sub btn_submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_submit.Click


        For Each ctl As Control In tab_breakdown.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
        For Each ctl As Control In tab_breakdown.Controls
            If TypeOf ctl Is ListBox Then
                DirectCast(ctl, ListBox).ClearSelected()
            End If
        Next


        If btnradio_payment_1.Checked = True Or btnradio_payment_2.Checked = True Or btnradio_payment_3.Checked = True Then
            Me.tab_pcOrder.TabPages(3).Enabled = True
            tab_pcOrder.SelectedTab = tab_breakdown

            tbx_staffid_breakdown.Text = tbx_staffid.Text
            tbxSummary_name.Text = tbx_firstname.Text & " " & tbx_surname.Text
            lstbxSummary_address.Items.Add(tbx_houseno.Text)
            lstbxSummary_address.Items.Add(tbx_street.Text)
            lstbxSummary_address.Items.Add(tbx_county.Text)
            lstbxSummary_address.Items.Add(tbx_postcode.Text)
            txtbxSummary_phone.Text = tbx_telephone.Text

            For X = 0 To 13
                lstbxSummary_orders.Items.Add(shoppinglist(X))
            Next

            txtbxSummary_mnthlypay.Text = tbx_monthlypayment.Text
            txtbxSummary_deposit.Text = tbx_deposit.Text
            txtbxSummary_interest.Text = tbx_interest.Text
            txtbxSummary_totCost.Text = tbx_totalcost.Text

            tab_pcOrder.TabPages(2).Enabled = False
        Else
            MsgBox("Please select Payment Option")
        End If

        '--------------------------DATABASE CONNECTION GENERATING A UNIQUE USER ID------------------------------------

        databseconnection()

        tbx_orderID.Text = tbx_firstname.Text.Substring(0, 1) & tbx_surname.Text & dsOrder.Tables("Orders").Rows(maxRowsOrder - 1).Item(0)
        
        'DATABASE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


    End Sub

    Private Sub btn_confirmOrder_Click(sender As Object, e As EventArgs) Handles btn_confirmOrder.Click
        Dim fileName As String = "U:\EDP_project\pc_OrderApp\pc_orderApp\TxtFile_Orders" + tbx_orderID.Text + ".txt"
        Dim objWriter As New System.IO.StreamWriter(fileName)

        MsgBox("Order " & tbx_orderID.Text & " Confirmed Successfully")

        '   WRITING TO A TEXT FILE
        objWriter.Write("Customer Name = " & tbxSummary_name.Text)
        objWriter.Write(Environment.NewLine)
        objWriter.Write(Environment.NewLine)
        objWriter.Write("ADDRESS")
        objWriter.Write(Environment.NewLine)
        For x = 0 To lstbxSummary_address.Items.Count - 1
            objWriter.Write(lstbxSummary_address.Items.Item(x))
            objWriter.Write(Environment.NewLine)
        Next
        objWriter.Write(Environment.NewLine)
        objWriter.Write("Contact = " & txtbxSummary_phone.Text)
        objWriter.Write(Environment.NewLine)
        objWriter.Write(Environment.NewLine)
        objWriter.Write("ITEMS ORDERED")
        objWriter.Write(Environment.NewLine)
        For x = 0 To lstbxSummary_orders.Items.Count - 1
            objWriter.Write(lstbxSummary_orders.Items.Item(x))
            objWriter.Write(Environment.NewLine)
        Next
        objWriter.Write(Environment.NewLine)
        objWriter.Write("Deposit = " & txtbxSummary_deposit.Text)
        objWriter.Write(Environment.NewLine)
        objWriter.Write("Interest = " & txtbxSummary_interest.Text)
        objWriter.Write(Environment.NewLine)
        objWriter.Write("Monthly Payment = " & txtbxSummary_mnthlypay.Text)
        objWriter.Write(Environment.NewLine)
        objWriter.Write("Total Cost = " & txtbxSummary_totCost.Text)
        objWriter.Close()
        ' END TEXT FILE WRITE FUNCTIONS

        cost = 0

        tab_pcOrder.TabPages.Remove(tab_customer_details)
        tab_pcOrder.TabPages.Remove(tab_orders)
        tab_pcOrder.TabPages.Remove(tab_payment)
        tab_pcOrder.TabPages.Remove(tab_breakdown)

        lstbxSummary_address.Items.Clear()
        lstbxSummary_orders.Items.Clear()
        tab_pcOrder.TabPages.Add(tab_login)

        For Each ctl As Control In tab_login.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
        tab_pcOrder.SelectedTab = tab_login

    End Sub

    '------------------ADMINISTRATOR TABS FUNCTIONS START-----------------------------------------
    '  -------------------------------------------------------------------------------------
    '         --------------------------------------------------------------------

    Private Sub btn_adminAccess_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminAccess.Click

        For Each ctl As Control In tab_administrator.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next

        'REMOVING LOG IN TAB
        tab_pcOrder.TabPages.Remove(tab_login)
        'DISPLAYING THE ADMINISTRATOR ACCESS TAB
        tab_pcOrder.TabPages.Add(tab_administrator)


    End Sub
    Private Sub btn_adminLogin_Click(sender As Object, e As EventArgs) Handles btn_adminLogin.Click

        If txtbx_adminID.Text = ds.Tables("StaffLogin").Rows(1).Item(1) And txtbx_adminPass.Text = ds.Tables("StaffLogin").Rows(1).Item(2) Then
            MsgBox("Access Guaranteed")
            tab_pcOrder.TabPages.Remove(tab_administrator)
            tab_pcOrder.TabPages.Add(Manager_TAB)
            btn_adminCancel.Enabled = False
            btn_adminCommit.Enabled = False
        Else
            MsgBox("Access Denied")
        End If

    End Sub


    'ADMINISTRATOR STAFF DETAILS CONTROL FORM
    
    Private Sub btn_adminNext_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminNext.Click

        If inc < maxRows - 1 Then
            inc = inc + 1
            tb_adminID.Text = ds.Tables("StaffLogin").Rows(inc).Item(1)
            tb_adminPass.Text = ds.Tables("StaffLogin").Rows(inc).Item(2)
            tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(inc).Item(3)
            tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(inc).Item(4)
            rbox_adminaddress.Clear()
            rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(inc).Item(5))
        Else
            MsgBox("No Futher Records to Display")
        End If

    End Sub
    Private Sub btn_adminPrevious_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminPrevious.Click

        If inc > 0 Then
            inc = inc - 1
            tb_adminID.Text = ds.Tables("StaffLogin").Rows(inc).Item(1)
            tb_adminPass.Text = ds.Tables("StaffLogin").Rows(inc).Item(2)
            tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(inc).Item(3)
            tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(inc).Item(4)
            rbox_adminaddress.Clear()
            rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(inc).Item(5))
        Else
            MsgBox("No Futher Records to Display")
        End If

    End Sub


    Private Sub btn_adminAddnew_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminAddnew.Click


        btn_adminCommit.Enabled = True
        btn_adminCancel.Enabled = True
        btn_adminAddnew.Enabled = False
        btn_adminDelete.Enabled = False
        btn_adminSearch.Enabled = False

        tb_adminID.Clear()
        tb_adminPass.Clear()
        tb_adminFirstName.Clear()
        tb_adminSurname.Clear()
        rbox_adminaddress.Clear()

        tb_adminID.ReadOnly = False
        tb_adminPass.ReadOnly = False
        tb_adminFirstName.ReadOnly = False
        tb_adminSurname.ReadOnly = False
        rbox_adminaddress.ReadOnly = False


    End Sub

    Private Sub btn_adminCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminCancel.Click

        btn_adminCommit.Enabled = False
        btn_adminAddnew.Enabled = True
        btn_adminDelete.Enabled = True
        btn_adminCancel.Enabled = False
        btn_adminSearch.Enabled = True


        tb_adminID.ReadOnly = True
        tb_adminPass.ReadOnly = True
        tb_adminFirstName.ReadOnly = True
        tb_adminSurname.ReadOnly = True
        rbox_adminaddress.ReadOnly = True

        inc = 0

        tb_adminID.Text = ds.Tables("StaffLogin").Rows(inc).Item(1)
        tb_adminPass.Text = ds.Tables("StaffLogin").Rows(inc).Item(2)
        tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(inc).Item(3)
        tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(inc).Item(4)
        rbox_adminaddress.Clear()
        rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(inc).Item(5))
    End Sub

    Private Sub btn_adminCommit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_adminCommit.Click


        If inc <> -1 Then
            Dim cb As New OleDb.OleDbCommandBuilder(da)
            Dim dsNewRow As DataRow

            dsNewRow = ds.Tables("StaffLogin").NewRow()

            dsNewRow.Item(1) = tb_adminID.Text
            dsNewRow.Item(2) = tb_adminPass.Text
            dsNewRow.Item(3) = tb_adminFirstName.Text
            dsNewRow.Item(4) = tb_adminSurname.Text
            dsNewRow.Item(5) = rbox_adminaddress.Text

            ds.Tables("StaffLogin").Rows.Add(dsNewRow)

            da.Update(ds, "StaffLogin")

            MsgBox("New record Added to the Database")

            'specify type of database being used
            dbProvide = "PROVIDER= Microsoft.jet.OLEDB.4.0;"
            'specify source db location
            dbSource = "Data Source = U:\EDP_project\pc_OrderApp\pc_orderApp\StaffLogin.mdb"

            'use connection object to create a connection string
            con.ConnectionString = dbProvide & dbSource

            'open connection to db
            con.Open()

            sql = "SELECT * FROM tblStaff"
            'use adapter to run sql statement on connection
            da = New OleDb.OleDbDataAdapter(sql, con)
            'use adapter to put result in dataset
            da.Fill(ds, "StaffLogin")

            maxRows = ds.Tables("StaffLogin").Rows.Count
            'MsgBox("Connection successfull")

            'close connection to db
            con.Close()


            btn_adminCommit.Enabled = False
            btn_adminAddnew.Enabled = True
            btn_adminDelete.Enabled = True
            btn_adminCancel.Enabled = False
            btn_adminSearch.Enabled = True


            tb_adminID.ReadOnly = True
            tb_adminPass.ReadOnly = True
            tb_adminFirstName.ReadOnly = True
            tb_adminSurname.ReadOnly = True
            rbox_adminaddress.ReadOnly = True



            tb_adminID.Text = ds.Tables("StaffLogin").Rows(0).Item(1)
            tb_adminPass.Text = ds.Tables("StaffLogin").Rows(0).Item(2)
            tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(0).Item(3)
            tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(0).Item(4)
            rbox_adminaddress.Clear()
            rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(0).Item(5))

        End If


    End Sub

    Private Sub btn_adminDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_adminDelete.Click

        If MessageBox.Show("Do you really want to Delete this Record?", "Delete",
             MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then

            MsgBox("Operation Cancelled")
            Exit Sub

        End If

        Dim cb As New OleDb.OleDbCommandBuilder(da)

        ds.Tables("StaffLogin").Rows(inc).Delete()
        maxRows = maxRows - 1

        inc = 0
        da.Update(ds, "StaffLogin")

        tb_adminID.Text = ds.Tables("StaffLogin").Rows(inc).Item(1)
        tb_adminPass.Text = ds.Tables("StaffLogin").Rows(inc).Item(2)
        tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(inc).Item(3)
        tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(inc).Item(4)
        rbox_adminaddress.Clear()
        rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(inc).Item(5))

    End Sub
    Private Sub btn_adminSearch_Click(sender As Object, e As EventArgs) Handles btn_adminSearch.Click
        Dim searchID As String = txb_adminID.Text

        For x = 0 To maxRows - 1
            If searchID = ds.Tables("StaffLogin").Rows(x).Item(1) Then
                tb_adminID.Text = ds.Tables("StaffLogin").Rows(x).Item(1)
                tb_adminPass.Text = ds.Tables("StaffLogin").Rows(x).Item(2)
                tb_adminFirstName.Text = ds.Tables("StaffLogin").Rows(x).Item(3)
                tb_adminSurname.Text = ds.Tables("StaffLogin").Rows(x).Item(4)
                rbox_adminaddress.Clear()
                rbox_adminaddress.AppendText(ds.Tables("StaffLogin").Rows(x).Item(5))

                txb_adminID.Clear()
            End If
        Next

    End Sub
   

    Private Sub btn_staffAccess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_staffAccess.Click

        'DISPLAYING THE ADMINISTRATOR ACCESS TAB
        tab_pcOrder.TabPages.Remove(tab_administrator)
        'REMOVING LOG IN TAB
        tab_pcOrder.TabPages.Add(tab_login)


    End Sub


    Private Sub btn_databaseAccess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_databaseAccess.Click
        Process.Start("explorer.exe", "U:\EDP_project\pc_OrderApp\pc_orderApp\StaffLogin.mdb")
    End Sub

    Private Sub btn_adminLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_adminLogout.Click

        tab_pcOrder.TabPages.Remove(Manager_TAB)

        'CLEARING TEXTBOXES
        For Each ctl As Control In tab_administrator.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next

        tab_pcOrder.TabPages.Add(tab_administrator)

    End Sub
    '---------------------END ADMINISTRATOR CONTROLS-----------------------
    '    ----------------------------------------------------------
    '           ----------------------------------------




    '----------------------------ASSISTANCE FUNCTIONS-----------------------
    Private Sub btn_custHelp_Click(sender As Object, e As EventArgs) Handles btn_custHelp.Click

        MsgBox("//PLEASE NOTE FOLLOWING ARE THE GENERAL HELP IN ORDER TO USE THE SYSTEM//" & Environment.NewLine &
               "Every box must be completed with the correct details;" & Environment.NewLine &
               "Once each box is completed press NEXT to proceed to the orders screen.")

    End Sub

    Private Sub btn_orderHelp_Click(sender As Object, e As EventArgs) Handles btn_orderHelp.Click

        MsgBox("//PLEASE NOTE FOLLOWING ARE THE GENERAL HELP IN ORDER TO USE THE SYSTEM//" & Environment.NewLine &
               "Select from the different drop down list the relevant Hardware;" & Environment.NewLine &
               "Once the items have been selected, press NEXT to proceed;" & Environment.NewLine &
               "!!NOTE!!: the system will not proceed if no items have been selected.")

    End Sub


    Private Sub btn_PaymentHelp_Click(sender As Object, e As EventArgs) Handles btn_PaymentHelp.Click

        MsgBox("//PLEASE NOTE FOLLOWING ARE THE GENERAL HELP IN ORDER TO USE THE SYSTEM//" & Environment.NewLine &
              "Slect a payment option and the system will calculate all the costs;" & Environment.NewLine &
              "Once agreed with the client press SUBMIT" & Environment.NewLine &
              "Pres CANCEL to cancel the order;" & Environment.NewLine &
              "Press BACK to amend the order;" & Environment.NewLine &
              "!!NOTE!!: the system will not proceed if no payment option have been selected")
    End Sub



    Private Sub btn_summaryHelp_Click(sender As Object, e As EventArgs) Handles btn_summaryHelp.Click

        MsgBox("//PLEASE NOTE FOLLOWING ARE THE GENERAL HELP IN ORDER TO USE THE SYSTEM//" & Environment.NewLine &
            "----Press CONFIRM ORDER to save the order to a text file----" & Environment.NewLine &
            "----THe system will hten automatically log you out.")

    End Sub

    Private Sub btn_adminfirst_Click(sender As Object, e As EventArgs)

    End Sub
End Class

'VERIFY EVERY TEXTBOX HAS A VALUE IN MANAGER TAB
'DEFAULT VALUES IN TEXTBOXES