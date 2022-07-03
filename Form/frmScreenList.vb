Imports BlackCoffeeLibrary
Imports LeaveFilingSystem
Imports LeaveFilingSystem.dsLeaveFiling
Imports LeaveFilingSystem.dsLeaveFilingTableAdapters
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class frmScreenList
    Private Shared IsSent As Boolean = False
    Private adpLeaveFiling As New LeaveFilingTableAdapter
    Private adpScreening As New ScreeningTableAdapter
    Private bsScreening As New BindingSource
    Private connection As New clsConnection
    Private dbJeonsoft As New SqlDbMethod(connection.JeonsoftConnection)
    Private dbLeaveFiling As New SqlDbMethod(connection.LocalConnection)
    Private dbMain As New Main

    Private devEmailAddress As String = String.Empty
    Private devEmailPassword As String = String.Empty
    Private dicSearchCriteria As New Dictionary(Of String, Integer)
    Private dsLeaveFiling As New dsLeaveFiling
    Private dtScreening As New ScreeningDataTable
    Private employeeCode As String = String.Empty
    Private employeeId As Integer = 0
    Private employeeName As String = String.Empty
    Private indexPosition As Integer = 0
    Private indexScroll As Integer = 0
    Private isDebug As Boolean = False
    Private isFilterByAbsentFrom As Boolean = False
    Private isFilterByDiagnosis As Boolean = False
    Private isFilterByEmployeeName As Boolean = False
    Private isFilterByReason As Boolean = False
    Private isFilterByScreenDate As Boolean = False
    Private pageCount As Integer
    Private pageIndex As Integer
    Private pageSize As Integer
    Private positionName As String = String.Empty
    Private senderEmailAddress As String = String.Empty
    Private senderEmailPassword As String = String.Empty
    Private totalCount As Integer
    Public Sub New(_employeeId As Integer, _employeeCode As String, _employeeName As String, _positionName As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        employeeId = _employeeId
        employeeCode = _employeeCode
        employeeName = _employeeName
        positionName = _positionName

        txtUsername.Text = StrConv(employeeName, VbStrConv.ProperCase) & " / " & positionName
        isDebug = SickLeaveScreening.My.Settings.IsDebug
    End Sub

    Public Sub RefreshList()
        If dgvList IsNot Nothing AndAlso dgvList.CurrentRow IsNot Nothing Then Me.Invoke(New Action(AddressOf GetScrollingIndex))
        pageIndex = 0
        BindPage()
        If dgvList IsNot Nothing AndAlso dgvList.CurrentRow IsNot Nothing Then Me.Invoke(New Action(AddressOf SetScrollingIndex))
    End Sub

    Public Sub SendApproverNotif(leaveFileId As Integer, approverId As Integer, leaveType As String, employeeName As String, department As String, leaveDate As String,
                                 reason As String)
        Try
            Dim client As New SmtpClient()
            Dim message As New MailMessage()
            Dim messageBody As String = "<font size=""3"" face=""Segoe UI"" color=""black"">" &
                                        "Good day! <br> <br> " &
                                        "New leave application for your approval. Please check the information below for your reference. <br> <br> " &
                                        "<table style=""font-size: 20px;font-family:Segoe UI""> " &
                                        "<tr><td style=""width:10px""></td><td>Leave File ID: </td><td style=""width:50px""></td><td>" & leaveFileId & "</td></tr>" &
                                        "<tr><td style=""width:10px""></td><td>Leave Type: </td><td style=""width:50px""></td><td>" & leaveType & "</td></tr>" &
                                        "<tr><td style=""width:10px""></td><td>Employee Name: </td><td style=""width:50px""></td><td>" & employeeName & "</td></tr>" &
                                        "<tr><td style=""width:10px""></td><td>Department/Section: </td><td style=""width:50px""></td><td>" & department & "</td></tr>" &
                                        "<tr><td style=""width:10px""></td><td>Date: </td><td style=""width:50px""></td><td>" & leaveDate & "</td></tr>" &
                                        "<tr><td style=""width:10px""></td><td>Reason: </td><td style=""width:50px""></td><td>" & reason & "</td></tr>" &
                                        "</table>" &
                                        "<br>" &
                                        "Please check on your Leave Application System." &
                                        "<br> <br>" &
                                        "If you have any concerns, please call IT (Local 232). Thank you." &
                                        "<br> <br>" &
                                        "<em>This is a system-generated email. Please do not reply.</em>"

            message.From = New MailAddress(senderEmailAddress, "NBC Leave Application")

            Dim prmApprover(0) As SqlParameter
            prmApprover(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
            prmApprover(0).Value = approverId

            Using reader As IDataReader = dbLeaveFiling.ExecuteReader("SELECT TRIM(NbcEmailAddress) AS NbcEmailAddress, TRIM(EmployeeName) AS EmployeeName " &
                                                                      "FROM dbo.Employee WHERE EmployeeId = @EmployeeId", CommandType.Text, prmApprover)

                While reader.Read
                    If Not reader.Item("NbcEmailAddress") Is DBNull.Value Then
                        If isDebug = True Then
                            message.Subject = "Leave Notification"
                            message.To.Add(devEmailAddress)
                        Else
                            message.Subject = "Leave Notification"
                            message.To.Add(reader.Item("NbcEmailAddress").ToString.Trim)
                        End If
                    Else
                        message.Subject = "No Company Email - " & reader.Item("EmployeeName").ToString.Trim & ""
                        message.To.Add(devEmailAddress)
                    End If
                End While
                reader.Close()
            End Using

            message.IsBodyHtml = True 'set email as html to attach hyperlink
            message.Body = messageBody

            client.Host = "smtp.gmail.com"
            client.Port = 587
            client.UseDefaultCredentials = False
            client.EnableSsl = True
            client.Credentials = New Net.NetworkCredential(senderEmailAddress, senderEmailPassword)

            Dim userState As String = "userState"
            AddHandler client.SendCompleted, AddressOf SendCompletedCallback

            client.SendAsync(message, userState)

            lblStatus.Visible = True
            lblStatus.Text = "Sending email, please wait......"
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SendDevNotif(employeeId As Integer, employeeName As String, leaveTypeId As Integer, leaveType As String,
                                      departmentId As Integer, departmentName As String, teamId As Integer, teamName As String,
                                      positionId As Integer, positionName As String)
        Try
            Dim client As New SmtpClient()
            Dim message As New MailMessage()
            Dim messageBody As String = "<font size=""3"" face=""Segoe UI"" color=""black"">" &
                                       "Good day! <br> <br> " &
                                       "New Automatic ECQ Leave Filing with no recipient. <br> <br> " &
                                       "<table style=""font-size: 20px;font-family:Segoe UI""> " &
                                       "<tr><td style=""width:10px""></td><td>Leave Type: </td><td style=""width:50px""></td><td>" & leaveType & "  (" & leaveTypeId & ")" & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Employee Name: </td><td style=""width:50px""></td><td>" & employeeName & "  (" & employeeId & ")" & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Department: </td><td style=""width:50px""></td><td>" & departmentName & "  (" &
                                       departmentId & ")" & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Team: </td><td style=""width:50px""></td><td>" & teamName & "  (" &
                                       teamId & ")" & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Position: </td><td style=""width:50px""></td><td>" & positionName & " (" &
                                       positionId & ")" & "</td></tr>" &
                                       "</table>" &
                                       "<br> <br>" &
                                       "<em>This is a system-generated email. Please do not reply.</em>"

            message.From = New MailAddress(senderEmailAddress, "NBC Leave Application")
            message.To.Add(devEmailAddress)

            message.Subject = "Leave Notification"
            message.IsBodyHtml = True 'set email as html to attach hyperlink
            message.Body = messageBody

            client.Host = "smtp.gmail.com"
            client.Port = 587
            client.UseDefaultCredentials = False
            client.EnableSsl = True
            client.Credentials = New Net.NetworkCredential(senderEmailAddress, senderEmailPassword)

            Dim userState As String = "userState"
            AddHandler client.SendCompleted, AddressOf SendCompletedCallback

            client.SendAsync(message, userState)

            lblStatus.Visible = True
            lblStatus.Text = "Sending email, please wait......"
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SendRequestorNotif(employeeId As Integer, screenDate As String, leaveTypeName As String, leaveDate As String, quantity As Integer, reason As String,
                                      diagnosis As String, isFitToWork As String)
        Try
            Dim client As New SmtpClient()
            Dim message As New MailMessage()
            Dim messageBody As String = "<font size=""3"" face=""Segoe UI"" color=""black"">" &
                                       "Good day! <br> <br> " &
                                       "Kindly apply your " & leaveTypeName & " to Leave Application System. Please check if the details below are correct. <br> <br> " &
                                       "<table style=""font-size: 20px;font-family:Segoe UI""> " &
                                       "<tr><td style=""width:10px""></td><td>Screen Date: </td><td style=""width:50px""></td><td>" & screenDate & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Leave Date(s): </td><td style=""width:50px""></td><td>" & leaveDate & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Quantity: </td><td style=""width:50px""></td><td>" & quantity & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Reason: </td><td style=""width:50px""></td><td>" & reason & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Diagnosis: </td><td style=""width:50px""></td><td>" & diagnosis & "</td></tr>" &
                                       "<tr><td style=""width:10px""></td><td>Fit To Work: </td><td style=""width:50px""></td><td>" & isFitToWork & "</td></tr>" &
                                       "</table>" &
                                       "<br>" &
                                       "If you have any concerns, please call IT (Local 232). Thank you." &
                                       "<br> <br>" &
                                       "<em>This is a system-generated email. Please do not reply.</em>"

            message.From = New MailAddress(senderEmailAddress, "NBC Leave Application")

            Dim prmRequestor(0) As SqlParameter
            prmRequestor(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
            prmRequestor(0).Value = employeeId

            Using reader As IDataReader = dbJeonsoft.ExecuteReader("SELECT TRIM(EmailAddress) AS EmailAddress, TRIM(Name) AS Name FROM dbo.tblEmployees WHERE Id = @EmployeeId",
                                                                  CommandType.Text, prmRequestor)

                While reader.Read
                    If Not reader.Item("EmailAddress") Is DBNull.Value Then
                        If isDebug = True Then
                            message.Subject = "Leave Notification"
                            message.To.Add(devEmailAddress)
                        Else
                            message.Subject = "Leave Notification"
                            message.To.Add(reader.Item("EmailAddress").ToString.Trim)
                        End If
                    Else
                        message.Subject = "No Personal Email - " & reader.Item("Name").ToString.Trim & ""
                        message.To.Add(devEmailAddress)
                    End If
                End While
                reader.Close()
            End Using

            message.IsBodyHtml = True 'set email as html to attach hyperlink
            message.Body = messageBody

            client.Host = "smtp.gmail.com"
            client.Port = 587
            client.UseDefaultCredentials = False
            client.EnableSsl = True
            client.Credentials = New Net.NetworkCredential(senderEmailAddress, senderEmailPassword)

            Dim userState As String = "userState"
            AddHandler client.SendCompleted, AddressOf SendCompletedCallback

            client.SendAsync(message, userState)

            lblStatus.Visible = True
            lblStatus.Text = "Sending email, please wait......"
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'prevent form resizing when double clicked the titlebar or dragged
    Protected Overloads Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_NCLBUTTONDBLCLK As Integer = 163 'define doubleclick event
        Const WM_NCLBUTTONDOWN As Integer = 161 'define leftbuttondown event
        Const WM_SYSCOMMAND As Integer = 274 'define move action
        Const HTCAPTION As Integer = 2 'define that the WM_NCLBUTTONDOWN is at titlebar
        Const SC_MOVE As Integer = 61456 'trap move action
        'disable moving titleBar
        If (m.Msg = WM_SYSCOMMAND) AndAlso (m.WParam.ToInt32() = SC_MOVE) Then
            Exit Sub
        End If
        'track whether clicked on title bar
        If (m.Msg = WM_NCLBUTTONDOWN) AndAlso (m.WParam.ToInt32() = HTCAPTION) Then
            Exit Sub
        End If
        'disable double click on title bar
        If (m.Msg = WM_NCLBUTTONDBLCLK) Then
            Exit Sub
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub BindingNavigatorMoveFirstItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorMoveFirstItem.Click
        pageIndex = 0
        BindPage()
    End Sub

    Private Sub BindingNavigatorMoveLastItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorMoveLastItem.Click
        pageIndex = pageCount - 1
        BindPage()
    End Sub

    Private Sub BindingNavigatorMoveNextItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorMoveNextItem.Click
        pageIndex += 1
        If pageIndex > pageCount - 1 Then
            pageIndex = pageCount - 1
        End If
        BindPage()
    End Sub

    Private Sub BindingNavigatorMovePreviousItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorMovePreviousItem.Click
        pageIndex -= 1
        If pageIndex < 0 Then
            pageIndex = 0
        End If
        BindPage()
    End Sub

    'can only press 0-9, delete, enter, backspace
    Private Sub BindingNavigatorPositionItem_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPageNumber.KeyPress
        If ((Asc(e.KeyChar) >= 48 AndAlso Asc(e.KeyChar) <= 57) OrElse Asc(e.KeyChar) = 8 OrElse Asc(e.KeyChar) = 13 OrElse Asc(e.KeyChar) = 127) Then
            e.Handled = False
            If Asc(e.KeyChar) = 13 Then
                Go()
            End If
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub BindPage()
        Try
            totalCount = 0

            If isFilterByScreenDate = True Then
                Me.adpScreening.FillByScreenDate(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount, dtpScreenDateFrom.Value.Date, dtpScreenDateTo.Value.Date)
            ElseIf isFilterByEmployeeName = True Then
                Me.adpScreening.FillByEmployeeName(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount, txtEmployeeName.Text.Trim)
            ElseIf isFilterByAbsentFrom = True Then
                Me.adpScreening.FillByAbsentFrom(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount, dtpAbsentFrom.Value.Date, dtpAbsentTo.Value.Date)
            ElseIf isFilterByReason = True Then
                Me.adpScreening.FillByReason(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount, txtReason.Text.Trim)
            ElseIf isFilterByDiagnosis = True Then
                Me.adpScreening.FillByDiagnosis(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount, txtDiagnosis.Text.Trim)
            Else
                Me.adpScreening.FillScreening(Me.dsLeaveFiling.Screening, pageIndex, pageSize, totalCount)
            End If

            Me.bsScreening.DataSource = Me.dsLeaveFiling
            Me.bsScreening.DataMember = dtScreening.TableName
            Me.bsScreening.ResetBindings(True)
            dgvList.AutoGenerateColumns = False
            Me.dgvList.DataSource = Me.bsScreening

            If totalCount Mod pageSize = 0 Then
                If totalCount = 0 Then
                    pageCount = (totalCount / pageSize) + 1
                Else
                    pageCount = totalCount / pageSize
                End If
            Else
                pageCount = Math.Truncate(totalCount / pageSize) + 1
            End If

            'current and total pages
            txtPageNumber.Text = pageIndex + 1
            txtTotalPageNumber.Text = " of " & CInt(pageCount) & " Page(s)"

            'enables pager
            txtPageNumber.Enabled = True
            txtTotalPageNumber.Enabled = True
            BindingNavigatorMoveFirstItem.Enabled = True
            BindingNavigatorMovePreviousItem.Enabled = True
            BindingNavigatorMoveNextItem.Enabled = True
            BindingNavigatorMoveLastItem.Enabled = True
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Try
            Using frmScreenEntry As New frmScreenEntry(employeeId)
                frmScreenEntry.ShowDialog(Me)
            End Using
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If dgvList.Rows.Count > 0 Then
                Dim screenId As Integer = CType(Me.bsScreening.Current, DataRowView).Item("ScreenId")
                Dim count As Integer = 0
                Dim leaveFileId As Integer = 0

                Dim prmCount(0) As SqlParameter
                prmCount(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                prmCount(0).Value = screenId

                count = dbLeaveFiling.ExecuteScalar("SELECT Count(LeaveFileId) FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, prmCount)

                If count > 0 Then
                    MessageBox.Show("Record was already used in the Leave Application System.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                Else
                    If MessageBox.Show("Delete this record?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        Me.bsScreening.RemoveCurrent()
                    End If
                End If

                Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                Me.dsLeaveFiling.AcceptChanges()
                RefreshList()
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDoctor_Click(sender As Object, e As EventArgs) Handles btnDoctor.Click
        Try
            Using frmDoctor As New frmDoctor()
                frmDoctor.ShowDialog(Me)
            End Using
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            If dgvList.Rows.Count > 0 Then
                Dim screenId As Integer = CType(Me.bsScreening.Current, DataRowView).Item("ScreenId")
                Using frmScreenEntry As New frmScreenEntry(employeeId, screenId)
                    frmScreenEntry.ShowDialog(Me)
                    If frmScreenEntry.DialogResult = Windows.Forms.DialogResult.OK Then
                        RefreshList()
                    End If
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Go()
    End Sub

    Private Sub btnLogOut_Click(sender As Object, e As EventArgs) Handles btnLogOut.Click
        Me.Hide()
        frmLogin.Show()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        RefreshList()
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Try
            Using frmHealthScreeningReport As New frmScreenReport()
                frmHealthScreeningReport.ShowDialog(Me)
            End Using
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            If cmbSearchCriteria.SelectedValue = 1 Then
                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = False

                dtpScreenDateFrom.Value = Date.Now
                dtpScreenDateTo.Value = Date.Now
                pageIndex = 0
                BindPage()

            ElseIf cmbSearchCriteria.SelectedValue = 2 Then
                txtEmployeeName.Clear()

                isFilterByScreenDate = False
                isFilterByEmployeeName = True
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = False

                pageIndex = 0
                BindPage()

            ElseIf cmbSearchCriteria.SelectedValue = 3 Then
                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = False

                dtpAbsentFrom.Value = Date.Now
                dtpAbsentTo.Value = Date.Now
                pageIndex = 0
                BindPage()

            ElseIf cmbSearchCriteria.SelectedValue = 4 Then
                txtReason.Clear()

                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = True
                isFilterByDiagnosis = False

                pageIndex = 0
                BindPage()

            ElseIf cmbSearchCriteria.SelectedValue = 5 Then
                txtDiagnosis.Clear()

                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = True

                pageIndex = 0
                BindPage()
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            If cmbSearchCriteria.SelectedValue = 1 Then
                If dtpScreenDateFrom.Value.Date > dtpScreenDateTo.Value.Date Then
                    MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                isFilterByScreenDate = True
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = False

            ElseIf cmbSearchCriteria.SelectedValue = 2 Then
                isFilterByScreenDate = False
                isFilterByEmployeeName = True
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = False

            ElseIf cmbSearchCriteria.SelectedValue = 3 Then
                If dtpAbsentFrom.Value.Date > dtpAbsentTo.Value.Date Then
                    MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = True
                isFilterByReason = False
                isFilterByDiagnosis = False

            ElseIf cmbSearchCriteria.SelectedValue = 4 Then
                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = True
                isFilterByDiagnosis = False

            ElseIf cmbSearchCriteria.SelectedValue = 5 Then
                isFilterByScreenDate = False
                isFilterByEmployeeName = False
                isFilterByAbsentFrom = False
                isFilterByReason = False
                isFilterByDiagnosis = True
            End If

            pageIndex = 0
            BindPage()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbSearchCriteria_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbSearchCriteria.SelectedValueChanged
        Try
            If cmbSearchCriteria.SelectedValue = 1 Then
                pnlScreenDate.Visible = True
                pnlEmployeeName.Visible = False
                pnlAbsentDate.Visible = False
                pnlReason.Visible = False
                pnlDiagnosis.Visible = False
                Me.ActiveControl = dtpScreenDateFrom

            ElseIf cmbSearchCriteria.SelectedValue = 2 Then
                pnlScreenDate.Visible = False
                pnlEmployeeName.Visible = True
                pnlAbsentDate.Visible = False
                pnlReason.Visible = False
                pnlDiagnosis.Visible = False
                Me.ActiveControl = txtEmployeeName

            ElseIf cmbSearchCriteria.SelectedValue = 3 Then
                pnlScreenDate.Visible = False
                pnlEmployeeName.Visible = False
                pnlAbsentDate.Visible = True
                pnlReason.Visible = False
                pnlDiagnosis.Visible = False
                Me.ActiveControl = dtpAbsentFrom

            ElseIf cmbSearchCriteria.SelectedValue = 4 Then
                pnlScreenDate.Visible = False
                pnlEmployeeName.Visible = False
                pnlAbsentDate.Visible = False
                pnlReason.Visible = True
                pnlDiagnosis.Visible = False
                Me.ActiveControl = txtReason

            ElseIf cmbSearchCriteria.SelectedValue = 5 Then
                pnlScreenDate.Visible = False
                pnlEmployeeName.Visible = False
                pnlAbsentDate.Visible = False
                pnlReason.Visible = False
                pnlDiagnosis.Visible = True
                Me.ActiveControl = txtDiagnosis

            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvList_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellDoubleClick
        btnEdit.PerformClick()
    End Sub

    Private Sub FillSearchCriteria()
        dicSearchCriteria.Add(" Screening Date", 1)
        dicSearchCriteria.Add(" Employee Name", 2)
        dicSearchCriteria.Add(" Absent Date", 3)
        dicSearchCriteria.Add(" Reason", 4)
        dicSearchCriteria.Add(" Diagnosis", 5)
        cmbSearchCriteria.DisplayMember = "Key"
        cmbSearchCriteria.ValueMember = "Value"
        cmbSearchCriteria.DataSource = New BindingSource(dicSearchCriteria, Nothing)
    End Sub

    Private Sub frmScreenList_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        dgvList.Dispose()
    End Sub

    Private Sub frmScreenList_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.F2
                e.Handled = True
                btnAdd.PerformClick()
            Case Keys.F3
                e.Handled = True
                btnEdit.PerformClick()
            Case Keys.F4
                e.Handled = True
                btnDelete.PerformClick()
            Case Keys.F5
                e.Handled = True
                RefreshList()
        End Select
    End Sub

    Private Sub frmScreenList_Load(sender As Object, e As EventArgs) Handles Me.Load
        AddHandler Me.SizeChanged, AddressOf Me_SizeChanged
        Me.MaximizeBox = False

        pageIndex = 0
        pageSize = 100
        BindPage()

        FillSearchCriteria()

        dbMain.EnableDoubleBuffered(dgvList)

        Me.dgvList.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Me.dgvList.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        GetEmailSettings(2)

        Me.ActiveControl = dgvList
    End Sub
    Private Sub GetEmailSettings(settingId As Integer)
        Try
            Dim prm(0) As SqlParameter
            prm(0) = New SqlParameter("@SettingId", SqlDbType.Int)
            prm(0).Value = settingId

            Using reader As IDataReader = dbLeaveFiling.ExecuteReader("SELECT * FROM dbo.Setting WHERE SettingId = @SettingId", CommandType.Text, prm)
                While reader.Read
                    senderEmailAddress = reader.Item("SenderEmail").ToString.Trim
                    senderEmailPassword = reader.Item("SenderEmailPassword").ToString.Trim
                    devEmailAddress = reader.Item("DevEmail").ToString.Trim
                    devEmailPassword = reader.Item("DevEmailPassword").ToString.Trim
                End While
                reader.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GetScrollingIndex()
        indexScroll = dgvList.FirstDisplayedCell.RowIndex
        indexPosition = dgvList.CurrentRow.Index
    End Sub

    Private Sub Go()
        Try
            If String.IsNullOrEmpty(txtPageNumber.Text) Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            If CInt(txtPageNumber.Text) > pageCount Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            If CInt(txtPageNumber.Text) = 0 Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            pageIndex = CInt(txtPageNumber.Text) - 1
            BindPage()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Function HideStatus() As Task(Of Boolean)
        Await Task.Delay(2000)
        lblStatus.Visible = False
        Return True
    End Function

    'set the window state to maximized without overlapping the taskbar
    Private Sub Me_SizeChanged(sender As Object, e As EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then
            Me.MaximizeBox = True

        ElseIf Me.WindowState = FormWindowState.Maximized Then
            Me.MaximizeBox = False
        End If
    End Sub
    Private Async Sub SendCompletedCallback(sender As Object, e As AsyncCompletedEventArgs)
        Try
            Dim token As String = CStr(e.UserState)

            If e.Cancelled Then
                lblStatus.Text = "Sending canceled."
            End If

            If e.Error IsNot Nothing Then
                lblStatus.Text = e.Error.ToString
            Else
                lblStatus.Text = "Email sent, thank you."
            End If

            Await HideStatus()

            IsSent = True
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetScrollingIndex()
        dgvList.FirstDisplayedScrollingRowIndex = indexScroll
        If dgvList.Rows.Count > indexPosition Then
            dgvList.Rows(indexPosition).Selected = True
        Else
            dgvList.Rows(indexPosition - 1).Selected = True
        End If
        Me.bsScreening.Position = dgvList.SelectedCells(0).RowIndex
    End Sub

End Class