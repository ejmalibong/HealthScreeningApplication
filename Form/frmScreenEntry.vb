Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports BlackCoffeeLibrary
Imports LeaveFilingSystem
Imports LeaveFilingSystem.dsLeaveFiling
Imports LeaveFilingSystem.dsLeaveFilingTableAdapters

Public Class frmScreenEntry
    Private connection As New clsConnection
    Private dbLeaveFiling As New SqlDbMethod(connection.LocalConnection)
    Private dbJeonsoft As New SqlDbMethod(connection.JeonsoftConnection)
    Private dbMain As New Main

    Private dsLeaveFiling As New dsLeaveFiling
    Private adpScreening As New ScreeningTableAdapter
    Private adpLeaveFiling As New LeaveFilingTableAdapter
    Private dtScreening As New ScreeningDataTable
    Private dtLeaveFiling As New LeaveFilingDataTable
    Private bsScreening As New BindingSource
    Private bsLeaveFiling As New BindingSource
    Private WithEvents screenDate As Binding
    Private WithEvents absentFrom As Binding
    Private WithEvents absentTo As Binding

    Private screenId As Integer = 0
    Private screenBy As Integer = 0 'doctor, nurse

    Private employeeId As Integer = 0
    Private teamId As Integer = 0
    Private departmentId As Integer = 0
    Private departmentName As String = String.Empty
    Private positionId As Integer = 0

    Private arrSplitted() As String 'value from scanner

    Private isDebug As Boolean = False

    Public Sub New(ByVal _screenBy As Integer, Optional ByVal _screenId As Integer = 0)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        screenBy = _screenBy
        screenId = _screenId
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        dbLeaveFiling.FillCmbWithCaption("SELECT LeaveTypeId, TRIM(LeaveTypeName) AS LeaveTypeName " & _
                                         "FROM LeaveType WHERE IsClinic = 1 " & _
                                         "ORDER BY TRIM(LeaveTypeName) ", _
                                         CommandType.Text, "LeaveTypeId", "LeaveTypeName", _
                                         cmbLeaveType, "< Select Leave Type >")

        If Not screenId = 0 Then
            Me.adpScreening.FillByScreenId(Me.dsLeaveFiling.Screening, screenId)
            Me.bsScreening.DataSource = Me.dsLeaveFiling
            Me.bsScreening.DataMember = dtScreening.TableName
            Me.bsScreening.Position = Me.bsScreening.Find("ScreenId", screenId)

            Me.Text = "Record No. " & screenId

            txtEmployeeScanId.Enabled = False
            If CType(Me.bsScreening.Current, DataRowView).Item("EmployeeId") Is DBNull.Value Then
                employeeId = 0
            Else
                employeeId = CType(Me.bsScreening.Current, DataRowView).Item("EmployeeId")
            End If

            txtEmployeeCode.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "EmployeeCode"))
            txtEmployeeName.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "EmployeeName"))
            screenDate = New Binding("Text", Me.bsScreening.Current, "ScreenDate")
            txtDate.DataBindings.Add(screenDate)
            absentFrom = New Binding("Text", Me.bsScreening.Current, "AbsentFrom")
            txtAbsentFrom.DataBindings.Add(absentFrom)
            absentTo = New Binding("Text", Me.bsScreening.Current, "AbsentTo")
            txtAbsentTo.DataBindings.Add(absentTo)
            txtReason.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "Reason"))
            txtDiagnosis.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "Diagnosis"))

            cmbLeaveType.DataBindings.Add(New Binding("SelectedValue", Me.bsScreening.Current, "LeaveTypeId"))

            If CType(Me.bsScreening.Current, DataRowView).Item("IsFitToWork") = True Then
                chkNotFtw.Checked = False
            Else
                chkNotFtw.Checked = True
            End If

            If txtEmployeeCode.Text.Trim.Substring(0, 3).ToUpper.Trim.Equals("FMB") Then
                txtEmployeeName.ReadOnly = False
            Else
                txtEmployeeName.ReadOnly = True
            End If

            Me.ActiveControl = txtDiagnosis
            txtDiagnosis.Select(txtDiagnosis.Text.Trim.Length, 0)
        Else
            ResetForm()
        End If
    End Sub

    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                e.Handled = True
                btnClear.PerformClick()

                screenId = 0
                employeeId = 0
                departmentId = 0
                teamId = 0
                positionId = 0

                txtEmployeeCode.Text = ""
                txtEmployeeName.Clear()
                txtDate.Text = ""
                txtAbsentFrom.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
                txtAbsentFrom.ValidatingType = GetType(System.DateTime)
                txtAbsentTo.ReadOnly = True
                txtAbsentTo.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
                txtAbsentTo.ValidatingType = GetType(System.DateTime)
                txtReason.Clear()
            Case Keys.F10
                e.Handled = True
                btnSave.PerformClick()
            Case Keys.F11
                e.Handled = True
                NotFitToWork()
            Case Keys.F12
                e.Handled = True
            Case Keys.Enter
                Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End Select
    End Sub

    Private Sub txtEmployeeScanId_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEmployeeScanId.KeyDown
        If e.KeyCode.Equals(Keys.Enter) Then
            e.Handled = True
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            arrSplitted = Split(txtEmployeeScanId.Text.Trim, " ", 2)
            GetEmployeeInformation(arrSplitted(0).ToString)
        End If
    End Sub

    Private Sub cmbLeaveType_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbLeaveType.SelectedValueChanged
        If cmbLeaveType.SelectedValue = 14 Then
            chkNotFtw.Enabled = False
            chkNotFtw.CheckState = CheckState.Checked
        Else
            chkNotFtw.Enabled = True
            chkNotFtw.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) AndAlso String.IsNullOrEmpty(txtEmployeeCode.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If cmbLeaveType.SelectedValue = 0 Then
                MessageBox.Show("Please select leave type.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            If String.IsNullOrEmpty(txtReason.Text.Trim) Then
                MessageBox.Show("Please indicate the reason.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtReason
                Return
            End If

            If String.IsNullOrEmpty(txtDiagnosis.Text.Trim) Then
                MessageBox.Show("Please indicate the diagnosis.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtDiagnosis
                Return
            End If

            If CDate(txtAbsentFrom.Text).Date > CDate(txtAbsentTo.Text).Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtAbsentFrom
                Return
            End If

            If cmbLeaveType.SelectedValue = 12 AndAlso Not (CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date)) Then
                MessageBox.Show("Half-day leave should have the same dates.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            SaveFitToWork(chkNotFtw.Checked)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If Not screenId = 0 Then
                Dim _rowScreening As ScreeningRow = Me.dsLeaveFiling.Screening.FindByScreenId(screenId)
                Dim _count As Integer = 0
                Dim _leaveFileId As Integer = 0

                Dim _prmCount(0) As SqlParameter
                _prmCount(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                _prmCount(0).Value = screenId

                _count = dbLeaveFiling.ExecuteScalar("SELECT Count(LeaveFileId) FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, _prmCount)

                If _count > 0 Then
                    MessageBox.Show("Screening record already used.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                Else
                    If MessageBox.Show("Delete this record?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        Me.bsScreening.RemoveCurrent()
                    End If
                End If

                If Me.dsLeaveFiling.HasChanges Then
                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    Me.dsLeaveFiling.AcceptChanges()
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                End If
            Else
                Me.ActiveControl = txtEmployeeScanId
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ResetForm()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dateBinding_Format(sender As Object, e As ConvertEventArgs) Handles screenDate.Format
        If Not e.Value Is DBNull.Value Then
            e.Value = Format(e.Value, "MMMM dd, yyyy  HH:mm")
        Else
            e.Value = dbLeaveFiling.GetServerDate.ToString("MMMM dd, yyyy  HH:mm")
        End If
    End Sub

    Private Sub maskedFrom_Format(sender As Object, e As ConvertEventArgs) Handles absentFrom.Format
        e.Value = Format(e.Value, "MM/dd/yyyy")
    End Sub

    Private Sub maskedTo_Format(sender As Object, e As ConvertEventArgs) Handles absentTo.Format
        e.Value = Format(e.Value, "MM/dd/yyyy")
    End Sub

    Private Sub txtEmployeeName_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtEmployeeName.Validating
        If String.IsNullOrEmpty(txtEmployeeName.Text.Trim) Then
            MessageBox.Show("Employee name is required.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    'validates input from masked textbox - it should be in MM/dd/yyyy format
    Private Sub txtAbsentFrom_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles txtAbsentFrom.TypeValidationCompleted
        If (Not e.IsValidInput) Then
            SendKeys.Send("{End}")
            MessageBox.Show("Please input date in Month/Day/Year format.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub txtAbsentTo_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles txtAbsentTo.TypeValidationCompleted
        If (Not e.IsValidInput) Then
            SendKeys.Send("{End}")
            MessageBox.Show("Please input date in Month/Day/Year format.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub lblNotFtw_Click(sender As Object, e As EventArgs) Handles lblNotFtw.Click
        If chkNotFtw.Enabled = True Then
            If chkNotFtw.CheckState = CheckState.Checked Then
                chkNotFtw.Checked = False
            Else
                chkNotFtw.Checked = True
            End If
        End If
    End Sub

#Region "Subroutines"

    Private Sub GetEmployeeInformation(ByVal _employeeCode As String)
        Try
            Dim _count As Integer = 0
            Dim _prmCount(0) As SqlParameter
            _prmCount(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
            _prmCount(0).Value = _employeeCode

            _count = dbJeonsoft.ExecuteScalar("SELECT Count(Id) FROM viwGroupEmployees WHERE EmployeeCode = @EmployeeCode AND Active = 1", CommandType.Text, _prmCount)

            cmbLeaveType.SelectedValue = 1
            cmbLeaveType.Enabled = True

            If _count > 0 Then 'direct employee
                Dim _prmReader(0) As SqlParameter
                _prmReader(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                _prmReader(0).Value = _employeeCode

                Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("RdEmployee", CommandType.StoredProcedure, _prmReader)

                While _reader.Read
                    employeeId = _reader.Item("Id")
                    departmentId = _reader.Item("DepartmentId")
                    teamId = _reader.Item("TeamId")
                    txtEmployeeCode.Text = _reader.Item("EmployeeCode").ToString.Trim
                    txtEmployeeName.Text = _reader("EmployeeName").ToString.Trim

                    If Not _reader.Item("TeamName") Is DBNull.Value Then
                        If _reader.Item("DepartmentName").ToString.Trim.Equals(_reader.Item("TeamName").ToString.Trim) Then
                            departmentName = _reader.Item("DepartmentName").ToString.Trim
                        Else
                            departmentName = _reader.Item("DepartmentName").ToString.Trim & " - " & _reader.Item("TeamName").ToString.Trim
                        End If
                        teamId = _reader.Item("TeamId")
                    Else
                        departmentName = _reader.Item("DepartmentName").ToString.Trim
                    End If

                    positionId = _reader.Item("PositionId")
                End While
                _reader.Close()

                txtEmployeeScanId.Clear()
                txtEmployeeScanId.Enabled = False
                txtEmployeeName.Enabled = True
                txtEmployeeName.ReadOnly = True
                txtDate.Text = Format(dbLeaveFiling.GetServerDate, "MMMM dd, yyyy HH:mm")
                txtAbsentFrom.Enabled = True
                txtAbsentFrom.ReadOnly = False
                txtAbsentTo.Enabled = True
                txtAbsentTo.ReadOnly = False
                txtReason.Enabled = True
                txtReason.ReadOnly = False
                txtDiagnosis.Enabled = True
                txtDiagnosis.ReadOnly = False
                chkNotFtw.Enabled = True
                txtReason.Focus()
            Else 'agency employee (fmb)
                If _employeeCode.Substring(0, 3).ToUpper.Trim.Equals("FMB") Then
                    employeeId = 0
                    txtEmployeeScanId.Clear()
                    txtEmployeeScanId.Enabled = False
                    txtEmployeeCode.Text = _employeeCode
                    txtEmployeeCode.Text = StrConv(txtEmployeeCode.Text.Trim, VbStrConv.Uppercase)
                    txtEmployeeName.Enabled = True
                    txtEmployeeName.ReadOnly = False
                    txtDate.Text = Format(dbLeaveFiling.GetServerDate, "MMMM dd, yyyy HH:mm")
                    txtAbsentFrom.Enabled = True
                    txtAbsentFrom.ReadOnly = False
                    txtAbsentTo.Enabled = True
                    txtAbsentTo.ReadOnly = False
                    txtReason.Enabled = True
                    txtReason.ReadOnly = False
                    txtDiagnosis.Enabled = True
                    txtDiagnosis.ReadOnly = False
                    chkNotFtw.Enabled = True
                    txtEmployeeName.Focus()
                Else
                    MessageBox.Show("Employee not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtEmployeeScanId.Focus()
                    Return
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ResetForm()
        screenId = 0
        employeeId = 0
        departmentId = 0
        teamId = 0
        positionId = 0

        Me.Text = "New Record"
        txtEmployeeScanId.Enabled = True
        txtEmployeeScanId.Clear()
        txtEmployeeCode.Text = ""
        txtEmployeeName.Clear()
        txtEmployeeName.Enabled = False
        txtDate.Text = ""
        cmbLeaveType.Enabled = False
        cmbLeaveType.SelectedValue = 0
        txtAbsentFrom.Enabled = False
        txtAbsentFrom.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
        txtAbsentFrom.ValidatingType = GetType(System.DateTime)
        txtAbsentTo.Enabled = False
        txtAbsentTo.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
        txtAbsentTo.ValidatingType = GetType(System.DateTime)
        txtReason.Clear()
        txtReason.Enabled = False
        txtDiagnosis.Clear()
        txtDiagnosis.Enabled = False
        chkNotFtw.Enabled = False
        chkNotFtw.CheckState = CheckState.Unchecked
        txtEmployeeScanId.Focus()
    End Sub

    Private Sub SaveFitToWork(ByVal _isUnfitToWork As Boolean)
        Try
            Dim _frmScreenList As frmScreenList = TryCast(Me.Owner, frmScreenList)

            If screenId = 0 Then
                Dim _newScreeningRow As ScreeningRow = Me.dsLeaveFiling.Screening.NewScreeningRow

                With _newScreeningRow
                    .ScreenDate = dbLeaveFiling.GetServerDate
                    .ScreenBy = screenBy

                    If employeeId = 0 Then
                        .SetEmployeeIdNull()
                    Else
                        .EmployeeId = employeeId
                    End If

                    .EmployeeCode = txtEmployeeCode.Text.Trim
                    .EmployeeName = txtEmployeeName.Text.Trim
                    .AbsentFrom = CDate(txtAbsentFrom.Text)
                    .AbsentTo = CDate(txtAbsentTo.Text)

                    If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                        .Quantity = 0.5
                        .LeaveTypeId = 11
                    Else
                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        .LeaveTypeId = cmbLeaveType.SelectedValue
                    End If

                    .Reason = txtReason.Text.Trim
                    .Diagnosis = txtDiagnosis.Text.Trim

                    If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                        .IsFitToWork = False
                    Else
                        If _isUnfitToWork = True Then
                            .IsFitToWork = False
                        Else
                            .IsFitToWork = True
                        End If
                    End If

                    If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) Then
                        .IsUsed = True
                    Else
                        .IsUsed = False
                    End If

                    .ModifiedBy = screenBy
                    .ModifiedDate = dbLeaveFiling.GetServerDate
                End With
                Me.dsLeaveFiling.Screening.AddScreeningRow(_newScreeningRow)
                Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                'ecq leave and ecq leave - quarantine
                If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) AndAlso employeeId <> 0 Then
                    Dim _newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                    With _newRowLeaveFiling
                        _newRowLeaveFiling.DateCreated = dbLeaveFiling.GetServerDate
                        _newRowLeaveFiling.ScreenId = _newScreeningRow.ScreenId
                        .EmployeeId = employeeId
                        .DepartmentId = departmentId
                        .TeamId = teamId
                        .StartDate = CDate(txtAbsentFrom.Text)
                        .EndDate = CDate(txtAbsentTo.Text)
                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        .Reason = txtReason.Text.Trim
                        .LeaveCredits = 0
                        .LeaveBalance = 0
                        .ClinicIsApproved = 1
                        .ClinicId = screenBy
                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                        .ClinicRemarks = txtDiagnosis.Text.Trim
                        .IsLateFiling = 1
                        .LeaveTypeId = cmbLeaveType.SelectedValue
                        .ModifiedBy = screenBy
                        .ModifiedDate = dbLeaveFiling.GetServerDate
                        .IsEncoded = 0

                        .SuperiorIsApproved1 = 0
                        .SetSuperiorApprovalDate1Null()
                        .SetSuperiorRemarks1Null()

                        .SuperiorIsApproved2 = 0
                        .SetSuperiorApprovalDate2Null()
                        .SetSuperiorRemarks2Null()

                        .ManagerIsApproved = 0
                        .SetManagerApprovalDateNull()
                        .SetManagerRemarksNull()

                        Dim _prm(2) As SqlParameter
                        _prm(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                        _prm(0).Value = departmentId
                        _prm(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                        _prm(1).Value = teamId
                        _prm(2) = New SqlParameter("PositionId", SqlDbType.Int)
                        _prm(2).Value = positionId

                        Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, _prm)
                        Dim _superiorId1 As Integer = 0
                        Dim _superiorId2 As Integer = 0
                        Dim _managerId As Integer = 0

                        While _reader.Read
                            If _reader.Item("SuperiorId1") Is DBNull.Value Then
                                .SetSuperiorId1Null()

                                If _reader.Item("SuperiorId2") Is DBNull.Value Then
                                    .SetSuperiorId2Null()

                                    If employeeId = _reader.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                        .RoutingStatusId = 3
                                        .ManagerId = 70 'dgm
                                    Else
                                        .RoutingStatusId = 3
                                        .ManagerId = _reader.Item("ManagerId")
                                    End If
                                Else
                                    If employeeId = _reader.Item("SuperiorId2") Then
                                        .RoutingStatusId = 3
                                        .SetSuperiorId2Null()
                                    Else
                                        .RoutingStatusId = 4
                                        .SuperiorId2 = _reader.Item("SuperiorId2")
                                    End If
                                End If

                            Else
                                If employeeId = _reader.Item("SuperiorId1") Then
                                    .RoutingStatusId = 4
                                    .SetSuperiorId1Null()
                                Else
                                    .RoutingStatusId = 5
                                    .SuperiorId1 = _reader.Item("SuperiorId1")
                                End If

                                If _reader.Item("SuperiorId2") Is DBNull.Value Then
                                    .SetSuperiorId2Null()
                                Else
                                    .SuperiorId2 = _reader.Item("SuperiorId2")
                                End If
                            End If

                            .ManagerId = _reader.Item("ManagerId")
                            _managerId = _reader.Item("ManagerId")
                        End While
                        _reader.Close()
                    End With
                    Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(_newRowLeaveFiling)
                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                    If isDebug = False Then
                        If _newRowLeaveFiling.RoutingStatusId = 3 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.ManagerId, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.ManagerId, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        ElseIf _newRowLeaveFiling.RoutingStatusId = 4 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.SuperiorId2, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.SuperiorId2, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        ElseIf _newRowLeaveFiling.RoutingStatusId = 5 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.SuperiorId1, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.SuperiorId1, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        End If
                    End If
                End If

                Me.dsLeaveFiling.AcceptChanges()
                _frmScreenList.RefreshList()
                ResetForm()

            Else
                Dim _rowScreening As ScreeningRow = Me.dsLeaveFiling.Screening.FindByScreenId(screenId)
                Dim _count As Integer = 0
                Dim _leaveFileId As Integer = 0

                Dim _prmCount(0) As SqlParameter
                _prmCount(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                _prmCount(0).Value = screenId

                _count = dbLeaveFiling.ExecuteScalar("SELECT Count(LeaveFileId) FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, _prmCount)

                If _count > 0 Then 'screening record already used
                    Dim _prmReader(0) As SqlParameter
                    _prmReader(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    _prmReader(0).Value = screenId

                    Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("SELECT LeaveFileId FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, _prmReader)

                    While _reader.Read
                        _leaveFileId = _reader.Item("LeaveFileId")
                    End While
                    _reader.Close()

                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, _leaveFileId)

                    Dim _rowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(_leaveFileId)

                    With _rowLeaveFiling
                        If .IsManagerApprovalDateNull = False Then 'leave is already approved/disapproved by the last approver
                            MessageBox.Show("Leave is already approved/disapproved by the last approver.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return

                        Else 'leave already encoded but not yet approved/disapproved by the last approver
                            .StartDate = CDate(txtAbsentFrom.Text)
                            .EndDate = CDate(txtAbsentTo.Text)

                            .LeaveTypeId = cmbLeaveType.SelectedValue

                            If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                                .Quantity = 0.5
                            Else
                                .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                            End If

                            .Reason = txtReason.Text.Trim
                            .ClinicRemarks = txtDiagnosis.Text.Trim

                            Dim _prmEmpCode(0) As SqlParameter
                            _prmEmpCode(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                            _prmEmpCode(0).Value = screenBy

                            Dim _reader2 As IDataReader = dbLeaveFiling.ExecuteReader("RdClinic", CommandType.StoredProcedure, _prmEmpCode)

                            While _reader2.Read
                                .ClinicId = _reader2.Item("Id")
                                .ModifiedBy = _reader2.Item("Id")
                            End While
                            _reader2.Close()

                            .ModifiedDate = dbLeaveFiling.GetServerDate

                            With _rowScreening
                                .ScreenBy = screenBy

                                If employeeId = 0 Then
                                    .SetEmployeeIdNull()
                                Else
                                    .EmployeeId = employeeId
                                End If

                                .EmployeeCode = txtEmployeeCode.Text.Trim
                                .EmployeeName = txtEmployeeName.Text.Trim
                                .AbsentFrom = CDate(txtAbsentFrom.Text)
                                .AbsentTo = CDate(txtAbsentTo.Text)

                                If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                                    .Quantity = 0.5
                                Else
                                    .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                End If

                                .LeaveTypeId = cmbLeaveType.SelectedValue
                                .Reason = txtReason.Text.Trim
                                .Diagnosis = txtDiagnosis.Text.Trim
                                .ModifiedBy = screenBy
                                .ModifiedDate = dbLeaveFiling.GetServerDate

                                If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                    .IsFitToWork = False
                                Else
                                    If _isUnfitToWork = True Then
                                        .IsFitToWork = False
                                    Else
                                        .IsFitToWork = True
                                    End If
                                End If
                            End With
                        End If
                    End With

                Else 'screening record not yet used
                    With _rowScreening
                        .ScreenBy = screenBy

                        If employeeId = 0 Then
                            .SetEmployeeIdNull()
                        Else
                            .EmployeeId = employeeId
                        End If

                        .EmployeeCode = txtEmployeeCode.Text.Trim
                        .EmployeeName = txtEmployeeName.Text.Trim
                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                        .AbsentTo = CDate(txtAbsentTo.Text)

                        If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                            .Quantity = 0.5
                        Else
                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        End If

                        .LeaveTypeId = cmbLeaveType.SelectedValue
                        .Reason = txtReason.Text.Trim
                        .Diagnosis = txtDiagnosis.Text.Trim
                        .ModifiedBy = screenBy
                        .ModifiedDate = dbLeaveFiling.GetServerDate

                        If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                            .IsFitToWork = False
                        Else
                            If _isUnfitToWork = True Then
                                .IsFitToWork = False
                            Else
                                .IsFitToWork = True
                            End If
                        End If
                    End With

                    If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) AndAlso employeeId <> 0 Then
                        Dim _newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                        Dim _countId As Integer = 0
                        Dim _prmCountId(0) As SqlParameter
                        _prmCountId(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                        _prmCountId(0).Value = txtEmployeeCode.Text.Trim

                        _countId = dbJeonsoft.ExecuteScalar("SELECT Count(Id) FROM viwGroupEmployees WHERE EmployeeCode = @EmployeeCode AND Active = 1", CommandType.Text, _prmCountId)

                        If _countId > 0 Then 'direct employee
                            Dim _prmReader(0) As SqlParameter
                            _prmReader(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                            _prmReader(0).Value = txtEmployeeCode.Text.Trim

                            Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("RdEmployee", CommandType.StoredProcedure, _prmReader)

                            While _reader.Read
                                departmentId = _reader.Item("DepartmentId")
                                teamId = _reader.Item("TeamId")
                                positionId = _reader.Item("PositionId")
                            End While
                            _reader.Close()
                        End If

                        Dim _prm(2) As SqlParameter
                        _prm(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                        _prm(0).Value = departmentId
                        _prm(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                        _prm(1).Value = teamId
                        _prm(2) = New SqlParameter("PositionId", SqlDbType.Int)
                        _prm(2).Value = positionId

                        Dim _readerExist As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, _prm)

                        With _newRowLeaveFiling
                            .DateCreated = dbLeaveFiling.GetServerDate
                            .ScreenId = screenId
                            .EmployeeId = employeeId
                            .DepartmentId = departmentId
                            .TeamId = teamId
                            .StartDate = CDate(txtAbsentFrom.Text)
                            .EndDate = CDate(txtAbsentTo.Text)
                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                            .Reason = txtReason.Text.Trim
                            .LeaveCredits = 0
                            .LeaveBalance = 0
                            .ClinicIsApproved = 1
                            .ClinicId = screenBy
                            .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                            .ClinicRemarks = txtDiagnosis.Text.Trim
                            .IsLateFiling = 1
                            .LeaveTypeId = cmbLeaveType.SelectedValue
                            .ModifiedBy = screenBy
                            .ModifiedDate = dbLeaveFiling.GetServerDate
                            .IsEncoded = 0

                            .SuperiorIsApproved1 = 0
                            .SetSuperiorApprovalDate1Null()
                            .SetSuperiorRemarks1Null()

                            .SuperiorIsApproved2 = 0
                            .SetSuperiorApprovalDate2Null()
                            .SetSuperiorRemarks2Null()

                            .ManagerIsApproved = 0
                            .SetManagerApprovalDateNull()
                            .SetManagerRemarksNull()

                            While _readerExist.Read
                                If _readerExist.Item("SuperiorId1") Is DBNull.Value Then
                                    .SetSuperiorId1Null()

                                    If _readerExist.Item("SuperiorId2") Is DBNull.Value Then
                                        .SetSuperiorId2Null()

                                        If employeeId = _readerExist.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                            .RoutingStatusId = 3
                                            .ManagerId = 70 'dgm
                                        Else
                                            .RoutingStatusId = 3
                                            .ManagerId = _readerExist.Item("ManagerId")
                                        End If
                                    Else
                                        If employeeId = _readerExist.Item("SuperiorId2") Then
                                            .RoutingStatusId = 3
                                            .SetSuperiorId2Null()
                                        Else
                                            .RoutingStatusId = 4
                                            .SuperiorId2 = _readerExist.Item("SuperiorId2")
                                        End If
                                    End If

                                Else
                                    If employeeId = _readerExist.Item("SuperiorId1") Then
                                        .SetSuperiorId1Null()

                                        If _readerExist.Item("SuperiorId2") Is DBNull.Value Then
                                            .SetSuperiorId2Null()

                                            If employeeId = _readerExist.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                                .RoutingStatusId = 3
                                                .ManagerId = 70 'dgm
                                            Else
                                                .RoutingStatusId = 3
                                                .ManagerId = _readerExist.Item("ManagerId")
                                            End If
                                        Else
                                            If employeeId = _readerExist.Item("SuperiorId2") Then
                                                .RoutingStatusId = 3
                                                .SetSuperiorId2Null()
                                            Else
                                                .RoutingStatusId = 4
                                                .SuperiorId2 = _readerExist.Item("SuperiorId2")
                                            End If
                                        End If

                                    Else
                                        .RoutingStatusId = 5
                                        .SuperiorId1 = _readerExist.Item("SuperiorId1")
                                        .ManagerId = _readerExist.Item("ManagerId")
                                    End If
                                End If
                            End While
                        End With
                        Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(_newRowLeaveFiling)
                        Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                        If isDebug = False Then
                            If _newRowLeaveFiling.RoutingStatusId = 3 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.ManagerId, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.ManagerId, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            ElseIf _newRowLeaveFiling.RoutingStatusId = 4 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.SuperiorId2, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.SuperiorId2, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            ElseIf _newRowLeaveFiling.RoutingStatusId = 5 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.SuperiorId1, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.SuperiorId1, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            End If
                        End If
                    End If
                End If

                Me.Validate()
                Me.bsScreening.EndEdit()
                Me.bsLeaveFiling.EndEdit()

                If Me.dsLeaveFiling.HasChanges Then
                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)
                    Me.dsLeaveFiling.AcceptChanges()
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                End If
            End If

            employeeId = 0
            departmentId = 0
            teamId = 0
            positionId = 0
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'tag employee as `unfit to work` using shortcut key (F11)
    Private Sub NotFitToWork()
        Try
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) AndAlso String.IsNullOrEmpty(txtEmployeeCode.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If cmbLeaveType.SelectedValue = 0 Then
                MessageBox.Show("Please select leave type.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            If String.IsNullOrEmpty(txtReason.Text.Trim) Then
                MessageBox.Show("Please indicate the reason.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtReason
                Return
            End If

            If String.IsNullOrEmpty(txtDiagnosis.Text.Trim) Then
                MessageBox.Show("Please indicate the diagnosis.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtDiagnosis
                Return
            End If

            If CDate(txtAbsentFrom.Text).Date > CDate(txtAbsentTo.Text).Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtAbsentFrom
                Return
            End If

            If cmbLeaveType.SelectedValue = 12 AndAlso Not (CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date)) Then
                MessageBox.Show("Half-day leave should have the same dates.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            Dim _frmScreenList As frmScreenList = TryCast(Me.Owner, frmScreenList)

            If screenId = 0 Then
                Dim _newScreeningRow As ScreeningRow = Me.dsLeaveFiling.Screening.NewScreeningRow

                With _newScreeningRow
                    .ScreenDate = dbLeaveFiling.GetServerDate
                    .ScreenBy = screenBy

                    If employeeId = 0 Then
                        .SetEmployeeIdNull()
                    Else
                        .EmployeeId = employeeId
                    End If

                    .EmployeeCode = txtEmployeeCode.Text.Trim
                    .EmployeeName = txtEmployeeName.Text.Trim
                    .AbsentFrom = CDate(txtAbsentFrom.Text)
                    .AbsentTo = CDate(txtAbsentTo.Text)

                    If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                        .Quantity = 0.5
                        .LeaveTypeId = 11
                    Else
                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        .LeaveTypeId = cmbLeaveType.SelectedValue
                    End If

                    .Reason = txtReason.Text.Trim
                    .Diagnosis = txtDiagnosis.Text.Trim

                    .IsFitToWork = False

                    If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) Then
                        .IsUsed = True
                    Else
                        .IsUsed = False
                    End If

                    .ModifiedBy = screenBy
                    .ModifiedDate = dbLeaveFiling.GetServerDate
                End With
                Me.dsLeaveFiling.Screening.AddScreeningRow(_newScreeningRow)
                Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) AndAlso employeeId <> 0 Then
                    Dim _newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                    With _newRowLeaveFiling
                        _newRowLeaveFiling.DateCreated = dbLeaveFiling.GetServerDate
                        _newRowLeaveFiling.ScreenId = _newScreeningRow.ScreenId
                        .EmployeeId = employeeId
                        .DepartmentId = departmentId
                        .TeamId = teamId
                        .StartDate = CDate(txtAbsentFrom.Text)
                        .EndDate = CDate(txtAbsentTo.Text)
                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        .Reason = txtReason.Text.Trim
                        .LeaveCredits = 0
                        .LeaveBalance = 0
                        .ClinicIsApproved = 1
                        .ClinicId = screenBy
                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                        .ClinicRemarks = txtDiagnosis.Text.Trim
                        .IsLateFiling = 1
                        .LeaveTypeId = cmbLeaveType.SelectedValue
                        .ModifiedBy = screenBy
                        .ModifiedDate = dbLeaveFiling.GetServerDate
                        .IsEncoded = 0

                        .SuperiorIsApproved1 = 0
                        .SetSuperiorApprovalDate1Null()
                        .SetSuperiorRemarks1Null()

                        .SuperiorIsApproved2 = 0
                        .SetSuperiorApprovalDate2Null()
                        .SetSuperiorRemarks2Null()

                        .ManagerIsApproved = 0
                        .SetManagerApprovalDateNull()
                        .SetManagerRemarksNull()

                        Dim _prm(2) As SqlParameter
                        _prm(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                        _prm(0).Value = departmentId
                        _prm(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                        _prm(1).Value = teamId
                        _prm(2) = New SqlParameter("PositionId", SqlDbType.Int)
                        _prm(2).Value = positionId

                        Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, _prm)

                        While _reader.Read
                            If _reader.Item("SuperiorId1") Is DBNull.Value Then
                                .SetSuperiorId1Null()

                                If _reader.Item("SuperiorId2") Is DBNull.Value Then
                                    .SetSuperiorId2Null()

                                    If employeeId = _reader.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                        .RoutingStatusId = 3
                                        .ManagerId = 70 'dgm
                                    Else
                                        .RoutingStatusId = 3
                                        .ManagerId = _reader.Item("ManagerId")
                                    End If
                                Else
                                    If employeeId = _reader.Item("SuperiorId2") Then
                                        .RoutingStatusId = 3
                                        .SetSuperiorId2Null()
                                    Else
                                        .RoutingStatusId = 4
                                        .SuperiorId2 = _reader.Item("SuperiorId2")
                                    End If
                                End If

                            Else
                                If employeeId = _reader.Item("SuperiorId1") Then
                                    .RoutingStatusId = 4
                                    .SetSuperiorId1Null()
                                Else
                                    .RoutingStatusId = 5
                                    .SuperiorId1 = _reader.Item("SuperiorId1")
                                End If

                                If _reader.Item("SuperiorId2") Is DBNull.Value Then
                                    .SetSuperiorId2Null()
                                Else
                                    .SuperiorId2 = _reader.Item("SuperiorId2")
                                End If
                            End If

                            .ManagerId = _reader.Item("ManagerId")
                        End While
                    End With
                    Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(_newRowLeaveFiling)
                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                    If isDebug = False Then
                        If _newRowLeaveFiling.RoutingStatusId = 3 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.ManagerId, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.ManagerId, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        ElseIf _newRowLeaveFiling.RoutingStatusId = 4 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.SuperiorId2, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.SuperiorId2, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        ElseIf _newRowLeaveFiling.RoutingStatusId = 5 Then
                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                            _newRowLeaveFiling.SuperiorId1, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            Else
                                _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                           _newRowLeaveFiling.SuperiorId1, _
                                                            cmbLeaveType.Text, _
                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                            departmentName, _
                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                            txtReason.Text.Trim)
                            End If
                        End If
                    End If
                End If

                Me.dsLeaveFiling.AcceptChanges()
                _frmScreenList.RefreshList()
                ResetForm()

            Else
                Dim _rowScreening As ScreeningRow = Me.dsLeaveFiling.Screening.FindByScreenId(screenId)
                Dim _count As Integer = 0
                Dim _leaveFileId As Integer = 0

                Dim _prmCount(0) As SqlParameter
                _prmCount(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                _prmCount(0).Value = screenId

                _count = dbLeaveFiling.ExecuteScalar("SELECT Count(LeaveFileId) FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, _prmCount)

                If _count > 0 Then 'screening record already used
                    Dim _prmReader(0) As SqlParameter
                    _prmReader(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    _prmReader(0).Value = screenId

                    Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("SELECT LeaveFileId FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, _prmReader)

                    While _reader.Read
                        _leaveFileId = _reader.Item("LeaveFileId")
                    End While
                    _reader.Close()

                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, _leaveFileId)

                    Dim _rowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(_leaveFileId)

                    With _rowLeaveFiling
                        If .IsManagerApprovalDateNull = False Then 'leave is already approved/disapproved by the last approver
                            MessageBox.Show("Leave is already approved/disapproved by the last approver.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return

                        Else 'leave already encoded but not yet approved/disapproved by the last approver
                            .StartDate = CDate(txtAbsentFrom.Text)
                            .EndDate = CDate(txtAbsentTo.Text)

                            .LeaveTypeId = cmbLeaveType.SelectedValue

                            If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                                .Quantity = 0.5
                            Else
                                .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                            End If

                            .Reason = txtReason.Text.Trim
                            .ClinicRemarks = txtDiagnosis.Text.Trim

                            Dim _prmEmpCode(0) As SqlParameter
                            _prmEmpCode(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                            _prmEmpCode(0).Value = screenBy

                            Dim _reader2 As IDataReader = dbLeaveFiling.ExecuteReader("RdClinic", CommandType.StoredProcedure, _prmEmpCode)

                            While _reader2.Read
                                .ClinicId = _reader2.Item("Id")
                                .ModifiedBy = _reader2.Item("Id")
                            End While
                            _reader2.Close()

                            .ModifiedDate = dbLeaveFiling.GetServerDate

                            With _rowScreening
                                .ScreenBy = screenBy

                                If employeeId = 0 Then
                                    .SetEmployeeIdNull()
                                Else
                                    .EmployeeId = employeeId
                                End If

                                .EmployeeCode = txtEmployeeCode.Text.Trim
                                .EmployeeName = txtEmployeeName.Text.Trim
                                .AbsentFrom = CDate(txtAbsentFrom.Text)
                                .AbsentTo = CDate(txtAbsentTo.Text)

                                If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                                    .Quantity = 0.5
                                Else
                                    .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                End If

                                .LeaveTypeId = cmbLeaveType.SelectedValue
                                .Reason = txtReason.Text.Trim
                                .Diagnosis = txtDiagnosis.Text.Trim
                                .ModifiedBy = screenBy
                                .ModifiedDate = dbLeaveFiling.GetServerDate

                                .IsFitToWork = False
                            End With
                        End If
                    End With

                Else 'screening record not yet used
                    With _rowScreening
                        .ScreenBy = screenBy

                        If employeeId = 0 Then
                            .SetEmployeeIdNull()
                        Else
                            .EmployeeId = employeeId
                        End If

                        .EmployeeCode = txtEmployeeCode.Text.Trim
                        .EmployeeName = txtEmployeeName.Text.Trim
                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                        .AbsentTo = CDate(txtAbsentTo.Text)

                        If cmbLeaveType.SelectedValue = 12 Then 'half-day leave
                            .Quantity = 0.5
                        Else
                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                        End If

                        .LeaveTypeId = cmbLeaveType.SelectedValue
                        .Reason = txtReason.Text.Trim
                        .Diagnosis = txtDiagnosis.Text.Trim
                        .ModifiedBy = screenBy
                        .ModifiedDate = dbLeaveFiling.GetServerDate

                        .IsFitToWork = False
                    End With

                    If (cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14) AndAlso employeeId <> 0 Then
                        Dim _newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                        Dim _countId As Integer = 0
                        Dim _prmCountId(0) As SqlParameter
                        _prmCountId(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                        _prmCountId(0).Value = txtEmployeeCode.Text.Trim

                        _countId = dbJeonsoft.ExecuteScalar("SELECT Count(Id) FROM viwGroupEmployees WHERE EmployeeCode = @EmployeeCode AND Active = 1", CommandType.Text, _prmCountId)

                        If _countId > 0 Then 'direct employee
                            Dim _prmReader(0) As SqlParameter
                            _prmReader(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                            _prmReader(0).Value = txtEmployeeCode.Text.Trim

                            Dim _reader As IDataReader = dbLeaveFiling.ExecuteReader("RdEmployee", CommandType.StoredProcedure, _prmReader)

                            While _reader.Read
                                departmentId = _reader.Item("DepartmentId")
                                teamId = _reader.Item("TeamId")
                                positionId = _reader.Item("PositionId")
                            End While
                            _reader.Close()
                        End If

                        Dim _prm(2) As SqlParameter
                        _prm(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                        _prm(0).Value = departmentId
                        _prm(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                        _prm(1).Value = teamId
                        _prm(2) = New SqlParameter("PositionId", SqlDbType.Int)
                        _prm(2).Value = positionId

                        Dim _readerExist As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, _prm)

                        With _newRowLeaveFiling
                            .DateCreated = dbLeaveFiling.GetServerDate
                            .ScreenId = screenId
                            .EmployeeId = employeeId
                            .DepartmentId = departmentId
                            .TeamId = teamId
                            .StartDate = CDate(txtAbsentFrom.Text)
                            .EndDate = CDate(txtAbsentTo.Text)
                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                            .Reason = txtReason.Text.Trim
                            .LeaveCredits = 0
                            .LeaveBalance = 0
                            .ClinicIsApproved = 1
                            .ClinicId = screenBy
                            .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                            .ClinicRemarks = txtDiagnosis.Text.Trim
                            .IsLateFiling = 1
                            .LeaveTypeId = cmbLeaveType.SelectedValue
                            .ModifiedBy = screenBy
                            .ModifiedDate = dbLeaveFiling.GetServerDate
                            .IsEncoded = 0

                            .SuperiorIsApproved1 = 0
                            .SetSuperiorApprovalDate1Null()
                            .SetSuperiorRemarks1Null()

                            .SuperiorIsApproved2 = 0
                            .SetSuperiorApprovalDate2Null()
                            .SetSuperiorRemarks2Null()

                            .ManagerIsApproved = 0
                            .SetManagerApprovalDateNull()
                            .SetManagerRemarksNull()

                            While _readerExist.Read
                                If _readerExist.Item("SuperiorId1") Is DBNull.Value Then
                                    .SetSuperiorId1Null()

                                    If _readerExist.Item("SuperiorId2") Is DBNull.Value Then
                                        .SetSuperiorId2Null()

                                        If employeeId = _readerExist.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                            .RoutingStatusId = 3
                                            .ManagerId = 70 'dgm
                                        Else
                                            .RoutingStatusId = 3
                                            .ManagerId = _readerExist.Item("ManagerId")
                                        End If
                                    Else
                                        If employeeId = _readerExist.Item("SuperiorId2") Then
                                            .RoutingStatusId = 3
                                            .SetSuperiorId2Null()
                                        Else
                                            .RoutingStatusId = 4
                                            .SuperiorId2 = _readerExist.Item("SuperiorId2")
                                        End If
                                    End If

                                Else
                                    If employeeId = _readerExist.Item("SuperiorId1") Then
                                        .SetSuperiorId1Null()

                                        If _readerExist.Item("SuperiorId2") Is DBNull.Value Then
                                            .SetSuperiorId2Null()

                                            If employeeId = _readerExist.Item("ManagerId") Then 'employee is a manager, set dgm as the approver 3
                                                .RoutingStatusId = 3
                                                .ManagerId = 70 'dgm
                                            Else
                                                .RoutingStatusId = 3
                                                .ManagerId = _readerExist.Item("ManagerId")
                                            End If
                                        Else
                                            If employeeId = _readerExist.Item("SuperiorId2") Then
                                                .RoutingStatusId = 3
                                                .SetSuperiorId2Null()
                                            Else
                                                .RoutingStatusId = 4
                                                .SuperiorId2 = _readerExist.Item("SuperiorId2")
                                            End If
                                        End If

                                    Else
                                        .RoutingStatusId = 5
                                        .SuperiorId1 = _readerExist.Item("SuperiorId1")
                                        .ManagerId = _readerExist.Item("ManagerId")
                                    End If
                                End If
                            End While
                        End With
                        Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(_newRowLeaveFiling)
                        Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                        If isDebug = False Then
                            If _newRowLeaveFiling.RoutingStatusId = 3 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.ManagerId, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.ManagerId, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            ElseIf _newRowLeaveFiling.RoutingStatusId = 4 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.SuperiorId2, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.SuperiorId2, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            ElseIf _newRowLeaveFiling.RoutingStatusId = 5 Then
                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                                _newRowLeaveFiling.SuperiorId1, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                Else
                                    _frmScreenList.SendEmailApprovers(_newRowLeaveFiling.LeaveFileId, _
                                                               _newRowLeaveFiling.SuperiorId1, _
                                                                cmbLeaveType.Text, _
                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), _
                                                                departmentName, _
                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " & CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"), _
                                                                txtReason.Text.Trim)
                                End If
                            End If
                        End If
                    End If
                End If

                Me.Validate()
                Me.bsScreening.EndEdit()
                Me.bsLeaveFiling.EndEdit()

                If Me.dsLeaveFiling.HasChanges Then
                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)
                    Me.dsLeaveFiling.AcceptChanges()
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                End If
            End If

            employeeId = 0
            departmentId = 0
            teamId = 0
            positionId = 0
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Functions"

    'set the absent date to the last working date - excluding sunday, company holidays and legal holidays
    Private Function GetLastWorkingDay(ByVal _date As DateTime) As Date
        Try
            _date = _date.AddDays(-1)
            While IsHoliday(_date) Or IsWeekend(_date)
                _date = _date.AddDays(-1)
            End While
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return _date
    End Function

    Private Function IsWeekend(ByVal _date As Date) As Boolean
        If _date.DayOfWeek.Equals(DayOfWeek.Sunday) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsHoliday(ByVal _date As Date) As Boolean
        Dim _count As Integer

        Try
            Dim _prmHoliday(0) As SqlParameter
            _prmHoliday(0) = New SqlParameter("@HolidayDate", SqlDbType.Date)
            _prmHoliday(0).Value = _date.ToShortDateString
            _count = 0
            _count = dbLeaveFiling.ExecuteScalar("SELECT COUNT(HolidayId) FROM dbo.Holiday WHERE HolidayDate = @HolidayDate", CommandType.Text, _prmHoliday)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If _count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    'get the total number of days from start date up to end date - excluding holidays and sundays
    Private Function GetTotalDays(ByVal _startDate As Date, ByVal _endDate As Date) As Integer
        Dim _count As Integer = 0

        Try
            If _startDate.Date.Equals(_endDate.Date) Then
                _count = 1
            Else
                For _i As Integer = 0 To (_endDate - _startDate).Days
                    If Not IsHoliday(_startDate) Then
                        If Not IsWeekend(_startDate) Then
                            _count += 1
                        End If
                    End If
                    _startDate = _startDate.AddDays(1)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return _count
    End Function

#End Region

#Region "UI"
    Private Sub txtEmployeeScanId_Enter(sender As Object, e As EventArgs) Handles txtEmployeeScanId.Enter
        lblEmployeeScanId.ForeColor = Color.White
        lblEmployeeScanId.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtEmployeeScanId_Leave(sender As Object, e As EventArgs) Handles txtEmployeeScanId.Leave
        lblEmployeeScanId.ForeColor = Color.Black
        lblEmployeeScanId.BackColor = SystemColors.Control
    End Sub

    Private Sub txtEmployeeName_Enter(sender As Object, e As EventArgs) Handles txtEmployeeName.Enter
        lblEmployeeName.ForeColor = Color.White
        lblEmployeeName.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtEmployeeName_Leave(sender As Object, e As EventArgs) Handles txtEmployeeName.Leave
        lblEmployeeName.ForeColor = Color.Black
        lblEmployeeName.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbLeaveType_Enter(sender As Object, e As EventArgs) Handles cmbLeaveType.Enter
        lblLeaveType.ForeColor = Color.White
        lblLeaveType.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbLeaveType_Leave(sender As Object, e As EventArgs) Handles cmbLeaveType.Leave
        lblLeaveType.ForeColor = Color.Black
        lblLeaveType.BackColor = SystemColors.Control
    End Sub

    Private Sub txtAbsentFrom_Enter(sender As Object, e As EventArgs) Handles txtAbsentFrom.Enter
        lblAbsentFrom.ForeColor = Color.White
        lblAbsentFrom.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtAbsentFrom_Leave(sender As Object, e As EventArgs) Handles txtAbsentFrom.Leave
        lblAbsentFrom.ForeColor = Color.Black
        lblAbsentFrom.BackColor = SystemColors.Control
    End Sub

    Private Sub txtAbsentTo_Enter(sender As Object, e As EventArgs) Handles txtAbsentTo.Enter
        lblAbsentTo.ForeColor = Color.White
        lblAbsentTo.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtAbsentTo_Leave(sender As Object, e As EventArgs) Handles txtAbsentTo.Leave
        lblAbsentTo.ForeColor = Color.Black
        lblAbsentTo.BackColor = SystemColors.Control
    End Sub

    Private Sub txtReason_Enter(sender As Object, e As EventArgs) Handles txtReason.Enter
        lblReason.ForeColor = Color.White
        lblReason.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtReason_Leave(sender As Object, e As EventArgs) Handles txtReason.Leave
        lblReason.ForeColor = Color.Black
        lblReason.BackColor = SystemColors.Control
    End Sub

    Private Sub txtDiagnosis_Enter(sender As Object, e As EventArgs) Handles txtDiagnosis.Enter
        lblDiagnosis.ForeColor = Color.White
        lblDiagnosis.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtDiagnosis_Leave(sender As Object, e As EventArgs) Handles txtDiagnosis.Leave
        lblDiagnosis.ForeColor = Color.Black
        lblDiagnosis.BackColor = SystemColors.Control
    End Sub

    Private Sub chkNotFtw_Enter(sender As Object, e As EventArgs) Handles chkNotFtw.Enter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub chkNotFtw_Leave(sender As Object, e As EventArgs) Handles chkNotFtw.Leave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub lblNotFtw_Enter(sender As Object, e As EventArgs) Handles lblNotFtw.Enter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub lblNotFtw_Leave(sender As Object, e As EventArgs) Handles lblNotFtw.Leave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub lblNotFtw_MouseEnter(sender As Object, e As EventArgs) Handles lblNotFtw.MouseEnter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub lblNotFtw_MouseLeave(sender As Object, e As EventArgs) Handles lblNotFtw.MouseLeave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub chkNotFtw_MouseEnter(sender As Object, e As EventArgs) Handles chkNotFtw.MouseEnter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub chkNotFtw_MouseLeave(sender As Object, e As EventArgs) Handles chkNotFtw.MouseLeave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub
#End Region

End Class